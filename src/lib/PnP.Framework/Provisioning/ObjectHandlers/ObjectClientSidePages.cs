using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.News.DataModel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PnP.Core.Services;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers.Extensions;
using PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using PnP.Framework.Provisioning.ObjectHandlers.Utilities;
using PnP.Framework.Utilities;
using PnP.Framework.Utilities.CanvasControl;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using PnPCore = PnP.Core.Model.SharePoint;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectClientSidePages : ObjectHandlerBase
    {
        private PnPContext pnpContext;
        private PnPCore.IPage dummyPage;
        private Func<string> getTemplateFolderName;
        private const string ContentTypeIdField = "ContentTypeId";
        private const string FileRefField = "FileRef";
        private const string SPSitePageFlagsField = "_SPSitePageFlags";
        private static readonly Guid MultilingualPagesFeature = new Guid("24611c05-ee19-45da-955f-6602264abaf8");
        private static readonly Guid MixedRealityFeature = new Guid("2ac9c540-6db4-4155-892c-3273957f1926");

        public override string Name
        {
            get { return "ClientSidePages"; }
        }

        public override string InternalName => "ClientSidePages";

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                pnpContext = PnPCoreSdk.Instance.GetPnPContext(web.Context as ClientContext);

                getTemplateFolderName = () =>
                {
                    if (null == dummyPage)
                    {
                        // GET /sites/2022-08-deleteme/_api/web/lists?$select=Id%2cTitle%2cBaseTemplate%2cEnableVersioning%2cEnableMinorVersions%2cEnableModeration%2cForceCheckout%2cEnableFolderCreation%2cListItemEntityTypeFullName%2cRootFolder%2fServerRelativeUrl%2cRootFolder%2fUniqueId%2cFields&$expand=RootFolder%2cRootFolder%2fProperties%2cFields&$filter=BaseTemplate+eq+119&$top=100 HTTP/1.1
                        dummyPage = pnpContext.Web.NewPage();
                    }
                    return dummyPage.GetTemplatesFolder();
                };

                web.EnsureProperties(w => w.ServerRelativeUrl);

                // determine pages library
                string pagesLibrary = "SitePages";

                List<string> preCreatedPages = new List<string>();

                // Ensure the needed languages are enabled on the site
                EnsureWebLanguages(web, template, scope);
                // Ensure spaces is enabled
                EnsureSpaces(web, template, scope);

                var currentPageIndex = 0;
                // var retrievedFilesCacheHeu = new Dictionary<string, Microsoft.SharePoint.Client.File>();
                // pre create the needed pages so we can fill the needed tokens which might be used later on when we put web parts on those pages
                foreach (var clientSidePage in template.ClientSidePages)
                {
                    // 1. POST ProcessQuery -> GetFileByServerRelativePath = Existenzprüfung
                    // 2. GET /sites/2022-08-deleteme/_api/web/lists?$select=Id%2cTitle%2cBaseTemplate%2cEnableVersioning%2cEnableMinorVersions%2cEnableModeration%2cForceCheckout%2cEnableFolderCreation%2cListItemEntityTypeFullName%2cRootFolder%2fServerRelativeUrl%2cRootFolder%2fUniqueId%2cFields&$expand=RootFolder%2cRootFolder%2fProperties%2cFields&$filter=BaseTemplate+eq+119&$top=100 HTTP/1.1
                    // GONE: 3. GET(404) GET /sites/2022-08-deleteme/_api/Web/getFileByServerRelativePath(decodedUrl='%2Fsites%2F2022-08-deleteme%2FSitePages%2Fparzival-blank%2B-15007745.aspx')?$select=ListId%2cServerRelativeUrl%2cUniqueId%2cListItemAllFields%2f*%2cListItemAllFields%2fId%2cListItemAllFields%2fParentList%2fId%2cListItemAllFields%2fParentList%2fFields%2fInternalName%2cListItemAllFields%2fParentList%2fFields%2fFieldTypeKind%2cListItemAllFields%2fParentList%2fFields%2fTypeAsString%2cListItemAllFields%2fParentList%2fFields%2fTitle%2cListItemAllFields%2fParentList%2fFields%2fId&$expand=ListItemAllFields%2cListItemAllFields%2fParentList%2cListItemAllFields%2fParentList%2fFields HTTP/1.1
                    // 4. POST POST /sites/2022-08-deleteme/_api/web/getFolderById('e5371ba7-51c1-4fde-9a60-478b1205eb6e')/files/AddTemplateFile(urlOfFile='%2Fsites%2F2022-08-deleteme%2FSitePages%2Fparzival-blank%2B-15007745.aspx',templateFileType=3) HTTP/1.1
                    // 5. GET GET /sites/2022-08-deleteme/_api/Web/getFileByServerRelativePath(decodedUrl='%2Fsites%2F2022-08-deleteme%2FSitePages%2Fparzival-blank%2B-15007745.aspx')?$select=ListId%2cServerRelativeUrl%2cUniqueId%2cListItemAllFields%2f*%2cListItemAllFields%2fId%2cListItemAllFields%2fParentList%2fId%2cListItemAllFields%2fParentList%2fFields%2fInternalName%2cListItemAllFields%2fParentList%2fFields%2fFieldTypeKind%2cListItemAllFields%2fParentList%2fFields%2fTypeAsString%2cListItemAllFields%2fParentList%2fFields%2fTitle%2cListItemAllFields%2fParentList%2fFields%2fId&$expand=ListItemAllFields%2cListItemAllFields%2fParentList%2cListItemAllFields%2fParentList%2fFields HTTP/1.1
                    // 6. POST ProcessQuery -> ListItemProperties setzen
                    // 7. POST ProcessQuery -> GetFileByServerRelativePath

                    // 2024-02-22: down from 7 to 6 calls
                    var (_, parserActions, _) = Task.Run(() => PagePreCreator.PreCreatePageAsync(web, clientSidePage, pagesLibrary, getTemplateFolderName, (title, message) => {
                        currentPageIndex++;
                        WriteSubProgress(title, message, currentPageIndex, template.ClientSidePages.Count);

                    })).GetAwaiter().GetResult();
                    foreach (var parserAction in parserActions)
                    {
                        parserAction(parser, web);
                    }

                    if (clientSidePage.Translations.Any())
                    {
                        //Pages.ClientSidePage page = null;
                        PnPCore.IPage page = null;
                        string pageName = DeterminePageName(parser, clientSidePage);
                        if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
                        {
                            // Get the existing template page
                            //page = web.LoadClientSidePage($"{Pages.ClientSidePage.GetTemplatesFolder(pagesLibraryList)}/{pageName}");
                            page = pnpContext.Web.LoadClientSidePage($"{getTemplateFolderName()}/{pageName}");
                        }
                        else
                        {
                            // Get the existing page
                            page = pnpContext.Web.LoadClientSidePage(pageName);
                        }

                        if (page != null)
                        {

                            //Pages.TranslationStatusCollection availableTranslations = page.Translations();
                            var availableTranslations = page.GetPageTranslations();

                            // Trigger the creation of the translated pages
                            //Pages.TranslationStatusCreationRequest tscr = new Pages.TranslationStatusCreationRequest();
                            PnPCore.PageTranslationOptions tscr = new PnPCore.PageTranslationOptions();
                            foreach (var translatedClientSidePage in clientSidePage.Translations)
                            {
                                //if (availableTranslations.Items.Where(p => p.Culture.Equals(new CultureInfo(translatedClientSidePage.LCID).Name, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault() == null)
                                if (availableTranslations.TranslatedLanguages.Where(p => p.Culture.Equals(new CultureInfo(translatedClientSidePage.LCID).Name, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault() == null)
                                {
                                    tscr.AddLanguage(translatedClientSidePage.LCID);
                                }
                            }

                            // Pages.TranslationStatusCollection translationResults = null;
                            PnPCore.IPageTranslationStatusCollection translationResults = null;
                            if (tscr.LanguageCodes != null && tscr.LanguageCodes.Count > 0)
                            {
                                //translationResults = page.GenerateTranslations(tscr);
                                translationResults = page.TranslatePages(tscr);
                            }

                            //IEnumerable<Pages.TranslationStatus> combinedTranslationResults = new List<Pages.TranslationStatus>();
                            IEnumerable<PnPCore.IPageTranslationStatus> combinedTranslationResults = new List<PnPCore.IPageTranslationStatus>();

                            // Translation results will contain all available pages when ran
                            //if (translationResults != null && translationResults.Items.Count > 0)
                            if (translationResults != null && translationResults.TranslatedLanguages.Count > 0)
                            {
                                //combinedTranslationResults = combinedTranslationResults.Union(translationResults.Items);
                                combinedTranslationResults = combinedTranslationResults.Union(translationResults.TranslatedLanguages);
                            }
                            // No new translations generated, so take what we got as available translations
                            //else if (availableTranslations != null && availableTranslations.Items.Count > 0)
                            else if (availableTranslations != null && availableTranslations.TranslatedLanguages.Count > 0)
                            {
                                //combinedTranslationResults = combinedTranslationResults.Union(availableTranslations.Items);
                                combinedTranslationResults = combinedTranslationResults.Union(availableTranslations.TranslatedLanguages);
                            }

                            foreach (var createdTranslation in combinedTranslationResults)
                            {
                                //string url = UrlUtility.Combine(web.ServerRelativeUrl, createdTranslation.Path.DecodedUrl);
                                string url = UrlUtility.Combine(web.ServerRelativeUrl, createdTranslation.Path);
                                preCreatedPages.Add(url);
                                // Load up page tokens for these translations
                                var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
                                web.Context.Load(file, f => f.UniqueId, f => f.ServerRelativePath);
                                web.Context.ExecuteQueryRetry();
                                //retrievedFilesCacheHeu[url] = file;

                                // Fill token
                                var pageUrlForToken = file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray());
                                parser.AddToken(new PageUniqueIdToken(web, pageUrlForToken, file.UniqueId));
                                parser.AddToken(new PageUniqueIdEncodedToken(web, pageUrlForToken, file.UniqueId));
                            }
                        }
                    }
                }

                var getFromFileCacheByUrl = (string url) =>
                {
                    // HEU: don't cache files... precreating pages might be called with another context from another thread; the part below needs to get fresh instances
                    return (Microsoft.SharePoint.Client.File)null;
                };

                currentPageIndex = 0;
                // Iterate over the pages and create/update them
                foreach (var clientSidePage in template.ClientSidePages)
                {
                    // 1. GET /sites/2022-08-deleteme/_api/web/lists?$select=Id%2cTitle%2cBaseTemplate%2cEnableVersioning%2cEnableMinorVersions%2cEnableModeration%2cForceCheckout%2cEnableFolderCreation%2cListItemEntityTypeFullName%2cRootFolder%2fServerRelativeUrl%2cRootFolder%2fUniqueId%2cFields&$expand=RootFolder%2cRootFolder%2fProperties%2cFields&$filter=BaseTemplate+eq+119&$top=100 HTTP/1.1
                    // 2. GET /sites/2022-08-deleteme/_api/Web/getFileByServerRelativePath(decodedUrl='%2Fsites%2F2022-08-deleteme%2FSitePages%2Fparzival-blank%2B-15007745.aspx')?$select=ListId%2cServerRelativeUrl%2cUniqueId%2cListItemAllFields%2f*%2cListItemAllFields%2fId%2cListItemAllFields%2fParentList%2fId%2cListItemAllFields%2fParentList%2fFields%2fInternalName%2cListItemAllFields%2fParentList%2fFields%2fFieldTypeKind%2cListItemAllFields%2fParentList%2fFields%2fTypeAsString%2cListItemAllFields%2fParentList%2fFields%2fTitle%2cListItemAllFields%2fParentList%2fFields%2fId&$expand=ListItemAllFields%2cListItemAllFields%2fParentList%2cListItemAllFields%2fParentList%2fFields HTTP/1.1
                    // 3. POST batch
                    // - POST https://heinrichulbricht.sharepoint.com/sites/2022-08-deleteme/_api/Web/Lists(guid'd0dce654-b0b6-4cf6-b305-a83544aa5e10')/RenderListDataAsStream HTTP/1.1
                    // CACHED: 4. POST /sites/2022-08-deleteme/_api/web/GetClientSideWebParts
                    // 5. GET /sites/2022-08-deleteme/_api/Web/getFileByServerRelativePath(decodedUrl='%2Fsites%2F2022-08-deleteme%2FSitePages%2Fparzival-blank%2B-15007745.aspx')?$select=ListId%2cServerRelativeUrl%2cUniqueId%2cListItemAllFields%2f*%2cListItemAllFields%2fId%2cListItemAllFields%2fParentList%2fId%2cListItemAllFields%2fParentList%2fFields%2fInternalName%2cListItemAllFields%2fParentList%2fFields%2fFieldTypeKind%2cListItemAllFields%2fParentList%2fFields%2fTypeAsString%2cListItemAllFields%2fParentList%2fFields%2fTitle%2cListItemAllFields%2fParentList%2fFields%2fId&$expand=ListItemAllFields%2cListItemAllFields%2fParentList%2cListItemAllFields%2fParentList%2fFields HTTP/1.1
                    // 6. POST /sites/2022-08-deleteme/_vti_bin/client.svc/ProcessQuery - Properties setzen
                    // 7. POST /sites/2022-08-deleteme/_vti_bin/client.svc/ProcessQuery - Properties holen
                    // 8. POST /sites/2022-08-deleteme/_vti_bin/client.svc/ProcessQuery - Properties holen
                    // 9. POST /sites/2022-08-deleteme/_vti_bin/client.svc/ProcessQuery - viele Properties holen...
                    // 10. POST /sites/2022-08-deleteme/_vti_bin/client.svc/ProcessQuery - Properties holen
                    // 11. POST /sites/2022-08-deleteme/_vti_bin/client.svc/ProcessQuery - Properties holen duplikate??
                    // 12. POST /sites/2022-08-deleteme/_vti_bin/client.svc/ProcessQuery - Properties setzen
                    // GONE: 13. POST /sites/2022-08-deleteme/_api/web/getFileById('a306f672-4d4f-4f0e-856d-76e2e9999154')/listitemallfields/SetCommentsDisabled HTTP/1.1
                    // 14. GET /sites/2022-08-deleteme/_api/Web/getFileByServerRelativePath(decodedUrl='%2Fsites%2F2022-08-deleteme%2FSitePages%2Fparzival-blank%2B-15007745.aspx')?$select=CheckOutType%2cListId%2cUniqueId HTTP/1.1
                    // 15. GET /sites/2022-08-deleteme/_api/web/lists?$select=Id%2cTitle%2cBaseTemplate%2cEnableVersioning%2cEnableMinorVersions%2cEnableModeration%2cForceCheckout%2cEnableFolderCreation%2cListItemEntityTypeFullName%2cRootFolder%2fServerRelativeUrl%2cRootFolder%2fUniqueId%2cFields&$expand=RootFolder%2cRootFolder%2fProperties%2cFields&$filter=BaseTemplate+eq+119&$top=100 HTTP/1.1

                    // 2024-02-22 down from 15 to 10 calls (starting with the second page as caching kicks in)
                    CreatePage(web, template, parser, scope, clientSidePage, pagesLibrary, getFromFileCacheByUrl, ref currentPageIndex, preCreatedPages);

                    if (clientSidePage.Translations.Any())
                    {
                        foreach (var translatedClientSidePage in clientSidePage.Translations)
                        {
                            CreatePage(web, template, parser, scope, translatedClientSidePage, pagesLibrary, getFromFileCacheByUrl, ref currentPageIndex, preCreatedPages);
                        }
                    }
                }
            }

            WriteMessage("Done processing Client Side Pages", ProvisioningMessageType.Completed);
            return parser;
        }

        private static void EnsureSpaces(Web web, ProvisioningTemplate template, PnPMonitoredScope scope)
        {
            var spacesPages = template.ClientSidePages.Where(p => p.Layout != null && p.Layout.Equals(PnPCore.PageLayoutType.Spaces.ToString(), StringComparison.InvariantCultureIgnoreCase));
            if (spacesPages.Any())
            {
                try
                {
                    // Enable the MUI feature
                    web.ActivateFeature(ObjectClientSidePages.MixedRealityFeature);
                }
                catch (Exception ex)
                {
                    scope.LogError($"Mixed reality feature could not be enabled: {ex.Message}");
                    throw;
                }
            }
        }

        private static void EnsureWebLanguages(Web web, ProvisioningTemplate template, PnPMonitoredScope scope)
        {
            List<int> neededLanguages = new List<int>();
            int neededSourceLanguage = 0;

            foreach (var page in template.ClientSidePages.Where(p => p.Translations.Any()))
            {
                if (neededSourceLanguage == 0)
                {
                    neededSourceLanguage = page.LCID > 0 ? page.LCID : (template.RegionalSettings != null ? template.RegionalSettings.LocaleId : 0);
                }
                else
                {
                    // Source language should be the same for all pages in the template
                    if (neededSourceLanguage != page.LCID)
                    {
                        string error = "The pages in this template are based upon multiple source languages while all pages in a site must have the same source language";
                        scope.LogError(error);
                        throw new Exception(error);
                    }
                }

                foreach (var translatedPage in page.Translations)
                {
                    if (!neededLanguages.Contains(translatedPage.LCID))
                    {
                        neededLanguages.Add(translatedPage.LCID);
                    }
                }
            }

            // No translations found, bail out
            if (neededLanguages.Count == 0)
            {
                return;
            }

            try
            {
                // Enable the MUI feature
                web.ActivateFeature(ObjectClientSidePages.MultilingualPagesFeature);
            }
            catch (Exception ex)
            {
                scope.LogError($"Multilingual pages feature could not be enabled: {ex.Message}");
                throw;
            }

            // Check the "source" language
            web.EnsureProperties(p => p.Language, p => p.IsMultilingual);
            int sourceLanguage = Convert.ToInt32(web.Language);
            if (sourceLanguage != neededSourceLanguage)
            {
                string error = $"The web has source language {sourceLanguage} while the template expects {neededSourceLanguage}";
                scope.LogError(error);
                throw new Exception(error);
            }

            // Ensure the needed languages are available on this site
            if (!web.IsMultilingual)
            {
                web.IsMultilingual = true;
                web.Context.Load(web, w => w.SupportedUILanguageIds);
                web.Update();
            }
            else
            {
                web.Context.Load(web, w => w.SupportedUILanguageIds);
            }
            web.Context.ExecuteQueryRetry();

            var supportedLanguages = web.SupportedUILanguageIds;
            bool languageAdded = false;
            foreach (var language in neededLanguages)
            {
                if (!supportedLanguages.Contains(language))
                {
                    web.AddSupportedUILanguage(language);
                    languageAdded = true;
                }
            }

            if (languageAdded)
            {
                web.Update();
                web.Context.ExecuteQueryRetry();
            }
        }

        // page components shouldn't change for a site during the course of what WikiTraccs does
        private static readonly Dictionary<string, IEnumerable<PnPCore.IPageComponent>> pageComponentsCache = new();
        // cache our content type ID; don't need to look that up for every page again
        private static readonly Dictionary<(string siteUrl, string contentTypeIdFromTemplate), string> contentTypeIdOnSitePagesListCache = new();

        private void CreatePage(Web web, ProvisioningTemplate template, TokenParser parser, PnPMonitoredScope scope, BaseClientSidePage clientSidePage, string pagesLibrary, Func<string, Microsoft.SharePoint.Client.File> getFileThatHasAlreadyBeenRetrievedForPage, ref int currentPageIndex, List<string> preCreatedPages)
        {
            string pageName = DeterminePageName(parser, clientSidePage);
            string url = $"{pagesLibrary}/{pageName}";

            if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
            {
                if (clientSidePage is TranslatedClientSidePage)
                {
                    url = $"{pagesLibrary}/{pageName}";
                }
                else
                {
                    //url = $"{pagesLibrary}/{Pages.ClientSidePage.GetTemplatesFolder(pagesLibraryList)}/{pageName}";
                    url = $"{pagesLibrary}/{getTemplateFolderName()}/{pageName}";
                }
            }

            // Write page level status messages, needed in case many pages are provisioned
            currentPageIndex++;
            int totalPages = 0;
            foreach (var p in template.ClientSidePages)
            {
                totalPages++;
                if (p.Translations.Any())
                {
                    totalPages += p.Translations.Count;
                }
            }
            WriteSubProgress("Provision ClientSidePage", pageName, currentPageIndex, totalPages);

            url = UrlUtility.Combine(web.ServerRelativeUrl, url);

            var exists = true;

            // see https://github.com/pnp/pnpframework/issues/724 about the broken page issue that sometimes creates pages missing basic field values like ClientSideApplicationId
            bool isBrokenPage = false;
            Microsoft.SharePoint.Client.File file = getFileThatHasAlreadyBeenRetrievedForPage(url);
            if (null == file)
            {
                try
                {
                    file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
                    web.Context.Load(file);
                    web.Context.ExecuteQueryRetry();
                }
                catch (ServerException ex)
                {
                    if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                    {
                        exists = false;
                    }
                }
            }

            var updateMetadataNotContent = clientSidePage.FieldValues.ContainsKey("WT_UpdatePageMetaDataNotContent");
            var skipCommentToggle = clientSidePage.FieldValues.ContainsKey("WT_SkipCommentToggle");
            PnPCore.IPage page = null;
            if (exists)
            {
                if (clientSidePage.Overwrite || preCreatedPages.Contains(url))
                {
                    if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
                    {
                        // Get the existing template page
                        if (clientSidePage is TranslatedClientSidePage)
                        {
                            page = pnpContext.Web.LoadClientSidePage($"{pageName}");
                        }
                        else
                        {
                            page = pnpContext.Web.LoadClientSidePage($"{getTemplateFolderName()}/{pageName}");
                        }
                    }
                    else
                    {
                        // Get the existing page
                        page = pnpContext.Web.LoadClientSidePage(pageName);
                    }
                    // 2025-01-15: need to set desired editor type because the loaded page information does not yet contain this information?
                    if (null != page)
                    {
                        page.EditorType = clientSidePage.EditorType;
                    }

                    // normally the page can be gotten when the file exists; but there seem to be rare cases of broken pages were basic field values are missing, see https://github.com/pnp/pnpframework/issues/724
                    // these broken pages are detected here with the hope to fix them by re-populating those fields (ultimately via Page.SaveAsync)
                    var isPotentiallyBrokenPage = exists && null != file && null == page;
                    if (isPotentiallyBrokenPage)
                    {
                        var item = file.ListItemAllFields;
                        // broken page; try to fix it
                        web.Context.Load(item);
                        web.Context.ExecuteQueryRetry();
                        var allCheckedFieldsAreMissingOrEmpty = true;
                        // arbitrarily chosen fields to take as indicator of broken pages - if those are empty the page is most likely broken
                        var checkForEmpty = new string[] { "ClientSideApplicationId", "Title", "PageLayoutType" };
                        foreach (var checkField in checkForEmpty)
                        {
                            if (item.FieldValues.ContainsKey(checkField) && !string.IsNullOrEmpty(item.FieldValues[checkField]?.ToString()))
                            {
                                allCheckedFieldsAreMissingOrEmpty = false;
                                break;
                            }
                        }
                        if (allCheckedFieldsAreMissingOrEmpty)
                        {
                            isBrokenPage = true;
                        }
                    }

                    // this check is only here for broken pages where aceesing the null page would throw
                    if (null != page)
                    {
                        // Clear the page
                        page.ClearPage();
                    }
                }
                else
                {
                    // HEU: adding metadata update mode; NOTE: below code is a duplicate from further down below
                    // ==========================================================================================
                    if (updateMetadataNotContent)
                    {
                        if (clientSidePage.FieldValues != null && clientSidePage.FieldValues.Any())
                        {
                            var pageListItem = file.ListItemAllFields;
                            // broken page; try to fix it
                            web.Context.Load(pageListItem);
                            web.Context.ExecuteQueryRetry();

                            // HEU: adjusted update logic depending on whether a page is new or not
                            // ==============================================================
                            var isNewlyCreatedPage = preCreatedPages.Contains(url);
                            // ==============================================================
                            ListItemUtilities.UpdateListItem(pageListItem, parser, clientSidePage.FieldValues, isNewlyCreatedPage ? ListItemUtilities.ListItemUpdateType.ForceUpdateOverwriteVersion : ListItemUtilities.ListItemUpdateType.UpdateOverwriteVersion);
                            return;
                        }
                        // ==========================================================================================
                    }
                    else
                    {
                        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ClientSidePages_NoOverWrite, pageName);
                        return;
                    }
                }
            }

            if (!exists || isBrokenPage)
            {
                // Create new client side page
                // OR for broken pages: re-populate all missing basic fields as well
                page = web.AddClientSidePage(clientSidePage.EditorType, pageName);
            }

            // Set page title
            string newTitle = parser.ParseString(clientSidePage.Title);
            if (page.PageTitle != newTitle)
            {
                page.PageTitle = newTitle;
            }

            // Set page layout
            if (!string.IsNullOrEmpty(clientSidePage.Layout))
            {
                page.LayoutType = (PnPCore.PageLayoutType)Enum.Parse(typeof(PnPCore.PageLayoutType), clientSidePage.Layout);
            }

            // Page Header
            if (clientSidePage.Header != null && page.LayoutType != PnPCore.PageLayoutType.Topic)
            {
                switch (clientSidePage.Header.Type)
                {
                    case ClientSidePageHeaderType.None:
                        {
                            page.RemovePageHeader();
                            break;
                        }
                    case ClientSidePageHeaderType.Default:
                        {
                            //Message ID: MC791596 / Roadmap ID: 386904 =>based on #1058 the PageTitle WebPart is not always in first section
                            if (clientSidePage.Sections.Any(s => s.Type == CanvasSectionType.OneColumnFullWidth && s.Controls.Any(c => c.Type == WebPartType.PageTitle)))
                            {
                                page.SetPageTitleWebPartPageHeader();
                            }
                            else
                            {
                                page.SetDefaultPageHeader();
                            }

                            // 2024-11-19 HEU ADDITION - also trigger setting the author fields for the page list item, which will be done by PnP.Core
                            // v===================================
                            if (clientSidePage.Header.AuthorByLineId > 0)
                            {
                                // note: those are set in the web part that is added via PnP template, don't set here
                                // page.PageHeader.Authors = clientSidePage.Header.Authors ?? "";
                                // page.PageHeader.AuthorByLine = clientSidePage.Header.AuthorByLine ?? "";

                                // BUT: this needs to be set so the page list item property is set by PnP Core (further down the road)
                                page.PageHeader.AuthorByLineId = clientSidePage.Header.AuthorByLineId;
                            }
                            // ^===================================
                            break;
                        }
                    case ClientSidePageHeaderType.Custom:
                        {
                            var serverRelativeImageUrl = parser.ParseString(clientSidePage.Header.ServerRelativeImageUrl);
                            if (clientSidePage.Header.TranslateX.HasValue && clientSidePage.Header.TranslateY.HasValue)
                            {
                                page.SetCustomPageHeader(serverRelativeImageUrl, clientSidePage.Header.TranslateX.Value, clientSidePage.Header.TranslateY.Value);
                            }
                            else
                            {
                                page.SetCustomPageHeader(serverRelativeImageUrl);
                            }

                            page.PageHeader.LayoutType = (PnPCore.PageHeaderLayoutType)Enum.Parse(typeof(PnPCore.PageHeaderLayoutType), clientSidePage.Header.LayoutType.ToString());
                            page.PageHeader.TextAlignment = (PnPCore.PageHeaderTitleAlignment)Enum.Parse(typeof(PnPCore.PageHeaderTitleAlignment), clientSidePage.Header.TextAlignment.ToString());
                            page.PageHeader.ShowTopicHeader = clientSidePage.Header.ShowTopicHeader;
                            page.PageHeader.ShowPublishDate = clientSidePage.Header.ShowPublishDate;
                            page.PageHeader.TopicHeader = parser.ParseString(clientSidePage.Header.TopicHeader);
                            page.PageHeader.AlternativeText = parser.ParseString(clientSidePage.Header.AlternativeText);
                            page.PageHeader.Authors = clientSidePage.Header.Authors;
                            page.PageHeader.AuthorByLine = clientSidePage.Header.AuthorByLine;
                            page.PageHeader.AuthorByLineId = clientSidePage.Header.AuthorByLineId;
                            break;
                        }
                }
            }

            if (!string.IsNullOrEmpty(clientSidePage.ThumbnailUrl))
            {
                page.ThumbnailUrl = parser.ParseString(clientSidePage.ThumbnailUrl);
            }

            // Add content on the page, not needed for repost pages
            if (page.LayoutType != PnPCore.PageLayoutType.RepostPage)
            {
                IEnumerable<PnPCore.IPageComponent> componentsToAdd;
                lock (pageComponentsCache)
                {
                    if (!pageComponentsCache.TryGetValue(page.PnPContext.Uri.ToString(), out componentsToAdd))
                    {
                        // Load existing available controls
                        componentsToAdd = page.AvailablePageComponents();
                        pageComponentsCache.Add(page.PnPContext.Uri.ToString(), componentsToAdd);
                    }
                }

                // if no section specified then add a default single column section
                if (!clientSidePage.Sections.Any())
                {
                    clientSidePage.Sections.Add(new CanvasSection() { Type = CanvasSectionType.OneColumn, Order = 10 });
                }

                int sectionCount = -1;
                // Apply the "layout" and content
                foreach (var section in clientSidePage.Sections)
                {
                    // Skip topic page header control section
                    if (section.Order == 999999)
                    {
                        continue;
                    }

                    sectionCount++;
                    switch (section.Type)
                    {
                        case CanvasSectionType.OneColumn:
                            page.AddSection(PnPCore.CanvasSectionTemplate.OneColumn, section.Order, (int)section.BackgroundEmphasis);
                            break;
                        case CanvasSectionType.OneColumnFullWidth:
                            page.AddSection(PnPCore.CanvasSectionTemplate.OneColumnFullWidth, section.Order, (int)section.BackgroundEmphasis);
                            break;
                        case CanvasSectionType.TwoColumn:
                            page.AddSection(PnPCore.CanvasSectionTemplate.TwoColumn, section.Order, (int)section.BackgroundEmphasis);
                            break;
                        case CanvasSectionType.ThreeColumn:
                            page.AddSection(PnPCore.CanvasSectionTemplate.ThreeColumn, section.Order, (int)section.BackgroundEmphasis);
                            break;
                        case CanvasSectionType.TwoColumnLeft:
                            page.AddSection(PnPCore.CanvasSectionTemplate.TwoColumnLeft, section.Order, (int)section.BackgroundEmphasis);
                            break;
                        case CanvasSectionType.TwoColumnRight:
                            page.AddSection(PnPCore.CanvasSectionTemplate.TwoColumnRight, section.Order, (int)section.BackgroundEmphasis);
                            break;
                        case CanvasSectionType.OneColumnVerticalSection:
                            page.AddSection(PnPCore.CanvasSectionTemplate.OneColumnVerticalSection, section.Order, (int)section.BackgroundEmphasis, (int)section.VerticalSectionEmphasis);
                            break;
                        case CanvasSectionType.TwoColumnVerticalSection:
                            page.AddSection(PnPCore.CanvasSectionTemplate.TwoColumnVerticalSection, section.Order, (int)section.BackgroundEmphasis, (int)section.VerticalSectionEmphasis);
                            break;
                        case CanvasSectionType.TwoColumnLeftVerticalSection:
                            page.AddSection(PnPCore.CanvasSectionTemplate.TwoColumnLeftVerticalSection, section.Order, (int)section.BackgroundEmphasis, (int)section.VerticalSectionEmphasis);
                            break;
                        case CanvasSectionType.TwoColumnRightVerticalSection:
                            page.AddSection(PnPCore.CanvasSectionTemplate.TwoColumnRightVerticalSection, section.Order, (int)section.BackgroundEmphasis, (int)section.VerticalSectionEmphasis);
                            break;
                        case CanvasSectionType.ThreeColumnVerticalSection:
                            page.AddSection(PnPCore.CanvasSectionTemplate.ThreeColumnVerticalSection, section.Order, (int)section.BackgroundEmphasis, (int)section.VerticalSectionEmphasis);
                            break;
                        default:
                            page.AddSection(PnPCore.CanvasSectionTemplate.OneColumn, section.Order, (int)section.BackgroundEmphasis);
                            break;
                    }

                    // key is web part instance ID, value ist the web part
                    var inlineImageWebParts = new Dictionary<string, PnPCore.IPageWebPart>();
                    // item1 is the text web part, item 2 is the referenced inline image web part instance ID
                    var textWebPartsReferencingInlineImages = new List<(PnPCore.IPageText, string)>();

                    // Configure collapsible section, if needed
                    if (section.Collapsible)
                    {
                        var targetSection = page.Sections[sectionCount];
                        targetSection.Collapsible = section.Collapsible;
                        targetSection.IsExpanded = section.IsExpanded;
                        targetSection.HeadingLevel = section.HeadingLevel;
                        targetSection.DisplayName = section.DisplayName;
                        targetSection.IconAlignment = (PnP.Core.Model.SharePoint.IconAlignment)Enum.Parse(
                            typeof(PnP.Core.Model.SharePoint.IconAlignment),
                            section.IconAlignment.ToString());
                        targetSection.ShowDividerLine = section.ShowDividerLine;
                    }

                    // Add controls to the section
                    if (section.Controls.Any())
                    {
                        // Safety measure: reset column order to 1 for columns marked with 0 or lower
                        foreach (var control in section.Controls.Where(p => p.Column <= 0).ToList())
                        {
                            control.Column = 1;
                        }

                        foreach (CanvasControl control in section.Controls)
                        {
                            PnPCore.IPageComponent baseControl = null;

                            // Is it a text control?
                            if (control.Type == WebPartType.Text)
                            {
                                var textControl = page.NewTextPart();

                                if (control.ControlProperties.Any())
                                {
                                    var textProperty = control.ControlProperties.First();
                                    textControl.Text = parser.ParseString(textProperty.Value);

                                    // search text for inline web part references
                                    var pattern = @"data-instance-id[ ]?=[ ]?\""([\w]{8}-[\w]{4}-[\w]{4}-[\w]{4}-[\w]{12})\""";
                                    var matches = Regex.Matches(textControl.Text, pattern);
                                    for (var i = 0; i < matches.Count; i++)
                                    {
                                        var match = matches[i];
                                        if (match.Success && match.Groups.Count == 2)
                                        {
                                            var inlineImageWebPartInstanceId = match.Groups[1].Value;
                                            textWebPartsReferencingInlineImages.Add((textControl, inlineImageWebPartInstanceId));
                                        }
                                    }
                                }
                                else
                                {
                                    if (!string.IsNullOrEmpty(control.JsonControlData))
                                    {
                                        var json = JsonConvert.DeserializeObject<Dictionary<string, string>>(control.JsonControlData);

                                        if (json.Count > 0)
                                        {
                                            textControl.Text = parser.ParseString(json.First().Value);
                                        }
                                    }
                                }
                                // Reduce column number by 1 due 0 start indexing
                                page.AddControl(textControl, page.Sections[sectionCount].Columns[control.Column - 1], control.Order);

                            }
                            // It is a web part
                            else
                            {
                                // apply token parsing on the web part properties
                                control.JsonControlData = parser.ParseString(control.JsonControlData);

                                // perform processing of web part properties (e.g. include listid property based list title property)
                                var webPartPostProcessor = CanvasControlPostProcessorFactory.Resolve(control);
                                webPartPostProcessor.Process(control, web.Context as ClientContext);

                                // Is a custom developed client side web part (3rd party)
                                if (control.Type == WebPartType.Custom)
                                {
                                    if (!string.IsNullOrEmpty(control.CustomWebPartName))
                                    {
                                        baseControl = componentsToAdd.FirstOrDefault(p => p.Name.Equals(control.CustomWebPartName, StringComparison.InvariantCultureIgnoreCase));
                                    }
                                    else if (control.ControlId != Guid.Empty)
                                    {
                                        baseControl = componentsToAdd.FirstOrDefault(p => p.Id.Equals($"{{{control.ControlId}}}", StringComparison.CurrentCultureIgnoreCase));

                                        if (baseControl == null)
                                        {
                                            baseControl = componentsToAdd.FirstOrDefault(p => p.Id.Equals(control.ControlId.ToString(), StringComparison.InvariantCultureIgnoreCase));
                                        }
                                    }
                                }
                                // Is an OOB client side web part (1st party)
                                else
                                {
                                    string webPartName = "";
                                    switch (control.Type)
                                    {
                                        case WebPartType.Image:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Image);
                                            break;
                                        case WebPartType.BingMap:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.BingMap);
                                            break;
                                        case WebPartType.Button:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Button);
                                            break;
                                        case WebPartType.CallToAction:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.CallToAction);
                                            break;
                                        case WebPartType.GroupCalendar:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.GroupCalendar);
                                            break;
                                        case WebPartType.News:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.News);
                                            break;
                                        case WebPartType.PowerBIReportEmbed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.PowerBIReportEmbed);
                                            break;
                                        case WebPartType.Sites:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Sites);
                                            break;
                                        case WebPartType.MicrosoftForms:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.MicrosoftForms);
                                            break;
                                        case WebPartType.ClientWebPart:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.ClientWebPart);
                                            break;
                                        case WebPartType.ContentEmbed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.ContentEmbed);
                                            break;
                                        case WebPartType.ContentRollup:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.ContentRollup);
                                            break;
                                        case WebPartType.DocumentEmbed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.DocumentEmbed);
                                            break;
                                        case WebPartType.Events:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Events);
                                            break;
                                        case WebPartType.Hero:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Hero);
                                            break;
                                        case WebPartType.ImageGallery:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.ImageGallery);
                                            break;
                                        case WebPartType.LinkPreview:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.LinkPreview);
                                            break;
                                        case WebPartType.List:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.List);
                                            break;
                                        case WebPartType.NewsFeed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.NewsFeed);
                                            break;
                                        case WebPartType.NewsReel:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.NewsReel);
                                            break;
                                        case WebPartType.PageTitle:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.PageTitle);
                                            break;
                                        case WebPartType.People:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.People);
                                            break;
                                        case WebPartType.QuickChart:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.QuickChart);
                                            break;
                                        case WebPartType.QuickLinks:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.QuickLinks);
                                            break;
                                        case WebPartType.SiteActivity:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SiteActivity);
                                            break;
                                        case WebPartType.VideoEmbed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.VideoEmbed);
                                            break;
                                        case WebPartType.YammerEmbed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.YammerEmbed);
                                            break;
                                        case WebPartType.CustomMessageRegion:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.CustomMessageRegion);
                                            break;
                                        case WebPartType.Divider:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Divider);
                                            break;
                                        case WebPartType.Spacer:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Spacer);
                                            break;
                                        case WebPartType.Kindle:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Kindle);
                                            break;
                                        case WebPartType.MyFeed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.MyFeed);
                                            break;
                                        case WebPartType.OrgChart:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.OrgChart);
                                            break;
                                        case WebPartType.SavedForLater:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SavedForLater);
                                            break;
                                        case WebPartType.Twitter:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Twitter);
                                            break;
                                        case WebPartType.WorldClock:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.WorldClock);
                                            break;
                                        case WebPartType.SpacesDocLib:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesDocLib);
                                            break;
                                        case WebPartType.SpacesFileViewer:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesFileViewer);
                                            break;
                                        case WebPartType.SpacesImageViewer:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesImageViewer);
                                            break;
                                        case WebPartType.SpacesModelViewer:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesModelViewer);
                                            break;
                                        case WebPartType.SpacesImageThreeSixty:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesImageThreeSixty);
                                            break;
                                        case WebPartType.SpacesVideoThreeSixty:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesVideoThreeSixty);
                                            break;
                                        case WebPartType.SpacesText2D:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesText2D);
                                            break;
                                        case WebPartType.SpacesVideoPlayer:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesVideoPlayer);
                                            break;
                                        case WebPartType.SpacesPeople:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesPeople);
                                            break;
                                    }

                                    baseControl = componentsToAdd.FirstOrDefault(p => p.Name.Equals(webPartName, StringComparison.InvariantCultureIgnoreCase));
                                }

                                if (baseControl != null)
                                {
                                    PnPCore.IPageWebPart myWebPart = page.NewWebPart(baseControl);
                                    myWebPart.Order = control.Order;
                                    //{
                                    //    Order = control.Order
                                    //};

                                    if (!string.IsNullOrEmpty(control.JsonControlData))
                                    {
                                        var json = JsonConvert.DeserializeObject<JObject>(control.JsonControlData);
                                        if (json["instanceId"] != null && json["instanceId"].Type != JTokenType.Null)
                                        {
                                            if (Guid.TryParse(json["instanceId"].Value<string>(), out Guid instanceId))
                                            {
                                                myWebPart.InstanceId = instanceId;
                                            }
                                        }
                                        if (json["title"] != null && json["title"].Type != JTokenType.Null)
                                        {
                                            myWebPart.Title = parser.ParseString(json["title"].Value<string>());
                                        }
                                        if (json["description"] != null && json["description"].Type != JTokenType.Null)
                                        {
                                            myWebPart.Description = parser.ParseString(json["description"].Value<string>());
                                        }

                                        // inline image; need to set RichTextEditorInstanceId after generating text web part
                                        if (json["properties"]?["isInlineImage"]?.Value<bool>() ?? false)
                                        {
                                            inlineImageWebParts.Add(myWebPart.InstanceId.ToString(), myWebPart);
                                        }
                                    }

                                    // Reduce column number by 1 due 0 start indexing
                                    page.AddControl(myWebPart, page.Sections[sectionCount].Columns[control.Column - 1], control.Order);

                                    // set properties using json string
                                    if (!string.IsNullOrEmpty(control.JsonControlData))
                                    {
                                        myWebPart.PropertiesJson = control.JsonControlData;
                                    }

                                    //CHECK:
                                    // set using property collection
                                    //if (control.ControlProperties.Any())
                                    //{
                                    //    // grab the "default" properties so we can deduct their types, needed to correctly apply the set properties
                                    //    var controlManifest = JObject.Parse(baseControl.Manifest);
                                    //    JToken controlProperties = null;
                                    //    if (controlManifest != null)
                                    //    {
                                    //        controlProperties = controlManifest.SelectToken("preconfiguredEntries[0].properties");
                                    //    }

                                    //    foreach (var property in control.ControlProperties)
                                    //    {
                                    //        Type propertyType = typeof(string);

                                    //        if (controlProperties != null)
                                    //        {
                                    //            var defaultProperty = controlProperties.SelectToken(property.Key, false);
                                    //            if (defaultProperty != null)
                                    //            {
                                    //                propertyType = Type.GetType($"System.{defaultProperty.Type}");

                                    //                if (propertyType == null)
                                    //                {
                                    //                    if (defaultProperty.Type.ToString().Equals("integer", StringComparison.InvariantCultureIgnoreCase))
                                    //                    {
                                    //                        propertyType = typeof(int);
                                    //                    }
                                    //                }
                                    //            }
                                    //        }

                                    //        myWebPart.Properties[property.Key] = JToken.FromObject(Convert.ChangeType(parser.ParseString(property.Value), propertyType));
                                    //    }
                                    //}
                                }
                                else
                                {

                                    PnPCore.IPageWebPart myWebPart = page.NewWebPart();
                                    myWebPart.Order = control.Order;


                                    if (!string.IsNullOrEmpty(control.JsonControlData))
                                    {
                                        var json = JsonConvert.DeserializeObject<JObject>(control.JsonControlData);
                                        if (json["instanceId"] != null && json["instanceId"].Type != JTokenType.Null)
                                        {
                                            if (Guid.TryParse(json["instanceId"].Value<string>(), out Guid instanceId))
                                            {
                                                myWebPart.InstanceId = instanceId;
                                            }
                                        }
                                        if (json["title"] != null && json["title"].Type != JTokenType.Null)
                                        {
                                            myWebPart.Title = parser.ParseString(json["title"].Value<string>());
                                        }
                                        if (json["description"] != null && json["description"].Type != JTokenType.Null)
                                        {
                                            myWebPart.Description = parser.ParseString(json["description"].Value<string>());
                                        }
                                        if (json["id"] != null && json["id"].Type != JTokenType.Null)
                                        {
                                            if (Guid.TryParse(json["id"].Value<string>(), out Guid webPartId))
                                            {
                                                var pageWebPartType = typeof(PnPCore.IPageWebPart).Assembly.GetType("PnP.Core.Model.SharePoint.PageWebPart");

                                                PropertyInfo propertyInfo = pageWebPartType.GetProperty("WebPartId");
                                                if (propertyInfo != null)
                                                {
                                                    propertyInfo.SetValue(myWebPart, json["id"].Value<string>());
                                                }
                                            }
                                        }
                                    }

                                    // Reduce column number by 1 due 0 start indexing
                                    page.AddControl(myWebPart, page.Sections[sectionCount].Columns[control.Column - 1], control.Order);

                                    // set properties using json string
                                    if (!string.IsNullOrEmpty(control.JsonControlData))
                                    {
                                        myWebPart.PropertiesJson = control.JsonControlData;
                                    }
                                }
                            }
                        }
                    }
                    foreach (var inlineImageWebPart in inlineImageWebParts)
                    {
                        var textWebPartReferencingInlineImage = textWebPartsReferencingInlineImages.Where(o => o.Item2.Equals(inlineImageWebPart.Key)).FirstOrDefault();
                        if (null != textWebPartReferencingInlineImage.Item1)
                        {
                            //var o = inlineImageWebPart.Value as object;
                            //o.GetType().GetProperty("RichTextEditorInstanceId").SetValue(o, textWebPartReferencingInlineImage.Item1.InstanceId.ToString(), null);
                            inlineImageWebPart.Value.RichTextEditorInstanceId = textWebPartReferencingInlineImage.Item1.InstanceId.ToString();
                        }
                    }
                }
            }

            // Handle the header controls in the topic pages
            if (page.LayoutType == PnPCore.PageLayoutType.Topic)
            {
                var headerControlSection = clientSidePage.Sections.FirstOrDefault(p => p.Order == 999999);
                if (headerControlSection != null)
                {
                    // Ensure there's at least one default section available
                    if (!page.Sections.Any())
                    {
                        page.AddSection(PnPCore.CanvasSectionTemplate.OneColumn, 0);
                    }

                    // Clear existing header controls as they'll be overwritten
                    page.HeaderControls.Clear();

                    // Load existing available controls
                    var componentsToAdd = page.AvailablePageComponents();

                    int order = 1;
                    foreach (var headerControl in headerControlSection.Controls)
                    {
                        PnPCore.IPageComponent baseControl = null;

                        // apply token parsing on the web part properties
                        headerControl.JsonControlData = parser.ParseString(headerControl.JsonControlData);

                        if (headerControl.Type == WebPartType.Custom)
                        {
                            // Find the base control installed to the current site
                            baseControl = componentsToAdd.FirstOrDefault(p => p.Id.Equals($"{{{headerControl.ControlId}}}", StringComparison.CurrentCultureIgnoreCase));
                            if (baseControl == null)
                            {
                                baseControl = componentsToAdd.FirstOrDefault(p => p.Id.Equals(headerControl.ControlId.ToString(), StringComparison.InvariantCultureIgnoreCase));
                            }

                            if (baseControl != null)
                            {
                                PnPCore.IPageWebPart myWebPart = page.NewWebPart(baseControl);

                                myWebPart.IsHeaderControl = true;

                                if (!string.IsNullOrEmpty(headerControl.JsonControlData))
                                {
                                    var json = JsonConvert.DeserializeObject<JObject>(headerControl.JsonControlData);
                                    if (json["instanceId"] != null && json["instanceId"].Type != JTokenType.Null)
                                    {
                                        if (Guid.TryParse(json["instanceId"].Value<string>(), out Guid instanceId))
                                        {
                                            myWebPart.InstanceId = instanceId;
                                        }
                                    }

                                    if (json["dataVersion"] != null && json["dataVersion"].Type != JTokenType.Null)
                                    {
                                        myWebPart.DataVersion = json["dataVersion"].Value<string>();
                                    }
                                }

                                // set properties using json string
                                if (!string.IsNullOrEmpty(headerControl.JsonControlData))
                                {
                                    myWebPart.PropertiesJson = headerControl.JsonControlData;
                                }

                                page.AddHeaderControl(myWebPart, order);
                                order++;
                            }
                            else
                            {
                                scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ClientSidePages_BaseControlNotFound, headerControl.ControlId, headerControl.CustomWebPartName);
                            }

                        }
                    }
                }
            }

            // Persist the page
            if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
            {
                page.SaveAsTemplate(pageName.Replace($"{getTemplateFolderName()}/", ""));
            }
            else
            {
                page.Save(pageName);
            }

            // Load the page list item
            var fileAfterSave = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
            web.Context.Load(fileAfterSave, p => p.ListItemAllFields);
            web.Context.ExecuteQueryRetry();

            // Update page content type
            bool isDirty = false;
            if (!string.IsNullOrEmpty(clientSidePage.ContentTypeID))
            {
                // ==============================================================================================
                // HEU optimization to prevent content types for a list to be loaded for every page provisioned
                // note: this has the assumption baked in that the list item has NOT yet the correct content type; this is specific to WikiTraccs and provisioning pages with custom content type derived from Site Page
                if (contentTypeIdOnSitePagesListCache.TryGetValue((page.PnPContext.Uri.ToString(), clientSidePage.ContentTypeID), out var bestMatchCt) && !string.IsNullOrEmpty(bestMatchCt))
                {
                    fileAfterSave.ListItemAllFields[ContentTypeIdField] = bestMatchCt;
                    isDirty = true;
                }
                else
                // ==============================================================================================
                {
                    ContentTypeId bestMatchCT = fileAfterSave.ListItemAllFields.ParentList.BestMatchContentTypeId(clientSidePage.ContentTypeID);
                    ContentTypeId currentCT = fileAfterSave.ListItemAllFields.FieldExistsAndUsed(ContentTypeIdField) ? ((ContentTypeId)fileAfterSave.ListItemAllFields[ContentTypeIdField]) : null;

                    if (currentCT == null)
                    {
                        fileAfterSave.ListItemAllFields[ContentTypeIdField] = bestMatchCT.StringValue;
                        contentTypeIdOnSitePagesListCache.Add((page.PnPContext.Uri.ToString(), clientSidePage.ContentTypeID), bestMatchCT.StringValue);
                        isDirty = true;
                    }
                    else if (currentCT != null && !currentCT.IsChildOf(bestMatchCT))
                    {
                        fileAfterSave.ListItemAllFields[ContentTypeIdField] = bestMatchCT.StringValue;
                        contentTypeIdOnSitePagesListCache.Add((page.PnPContext.Uri.ToString(), clientSidePage.ContentTypeID), bestMatchCT.StringValue);
                        isDirty = true;
                    }
                }
            }

            if (clientSidePage.PromoteAsTemplate && page.LayoutType == PnPCore.PageLayoutType.Article)
            {
                // Choice field, currently there's only one value possible and that's Template
                fileAfterSave.ListItemAllFields[SPSitePageFlagsField] = ";#Template;#";
                isDirty = true;
            }

            if (isDirty)
            {
                if (exists)
                {
                    fileAfterSave.ListItemAllFields.SystemUpdate();
                }
                else
                {
                    fileAfterSave.ListItemAllFields.UpdateOverwriteVersion();
                }
                web.Context.Load(fileAfterSave.ListItemAllFields);
                web.Context.ExecuteQueryRetry();
            }

            bool? isModerationEnabled = null;
            bool? isMinorVersionEnabled = null;

            if (clientSidePage.FieldValues != null && clientSidePage.FieldValues.Any())
            {
                // HEU: adjusted update logic depending on whether a page is new or not
                // ==============================================================
                var isNewlyCreatedPage = preCreatedPages.Contains(url);
                // ==============================================================
                ListItemUtilities.UpdateListItem(fileAfterSave.ListItemAllFields, parser, clientSidePage.FieldValues, isNewlyCreatedPage ? ListItemUtilities.ListItemUpdateType.ForceUpdateOverwriteVersion : ListItemUtilities.ListItemUpdateType.UpdateOverwriteVersion);
                isModerationEnabled = fileAfterSave.ListItemAllFields.ParentList.EnableModeration;
                isMinorVersionEnabled = fileAfterSave.ListItemAllFields.ParentList.EnableMinorVersions;
            }

            // Set page property bag values
            if (clientSidePage.Properties != null && clientSidePage.Properties.Any())
            {
                string pageFilePath = fileAfterSave.ListItemAllFields[FileRefField].ToString();
                var pageFile = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(pageFilePath));
                web.Context.Load(pageFile, p => p.Properties);

                foreach (var pageProperty in clientSidePage.Properties)
                {
                    if (!string.IsNullOrEmpty(pageProperty.Key))
                    {
                        pageFile.Properties[pageProperty.Key] = pageProperty.Value;
                    }
                }

                pageFile.Update();
                web.Context.Load(fileAfterSave.ListItemAllFields);
                web.Context.ExecuteQueryRetry();
            }

            if (page.LayoutType != PnPCore.PageLayoutType.SingleWebPartAppPage)
            {
                // Set commenting, ignore on pages of the type Home or page templates
                if (page.LayoutType != PnPCore.PageLayoutType.Home && !clientSidePage.PromoteAsTemplate)
                {
                    // Make it a news page if requested
                    if (clientSidePage.PromoteAsNewsArticle)
                    {
                        page.PromoteAsNewsArticle();
                    }
                }

                // HEU: save one call
                if (!skipCommentToggle)
                {
                    if (page.LayoutType != PnPCore.PageLayoutType.RepostPage)
                    {
                        if (clientSidePage.EnableComments)
                        {
                            page.EnableComments();
                        }
                        else
                        {
                            page.DisableComments();
                        }
                    }
                }
            }

            // Publish page, page templates cannot be published
            if (clientSidePage.Publish && !clientSidePage.PromoteAsTemplate)
            {
                page.PublishAsync(null, isMinorVersionEnabled, isModerationEnabled).GetAwaiter().GetResult();
            }

            // Set any security on the page
            if (clientSidePage.Security != null && clientSidePage.Security.RoleAssignments.Count != 0)
            {
                web.Context.Load(fileAfterSave.ListItemAllFields);
                web.Context.ExecuteQueryRetry();
                fileAfterSave.ListItemAllFields.SetSecurity(parser, clientSidePage.Security, WriteMessage);
            }
        }

#nullable enable
        public static string DeterminePageName(TokenParser? parser, BaseClientSidePage clientSidePage)
        {
            string pageName;
            if (clientSidePage is ClientSidePage csp)
            {
                if (clientSidePage.PromoteAsTemplate)
                {
                    pageName = parser?.ParseString(csp.PageName) ?? csp.PageName;
                    pageName = $"{System.IO.Path.GetFileNameWithoutExtension(pageName)}.aspx";
                }
                else
                {
                    var parsedPageName = parser?.ParseString(csp.PageName) ?? csp.PageName;
                    var pageNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(parsedPageName);
                    var pageFolder = System.IO.Path.GetDirectoryName(parsedPageName);
                    if (!string.IsNullOrEmpty(pageFolder))
                    {
                        pageFolder += "/";
                    }

                    pageName = $"{pageFolder}{pageNameWithoutExtension}.aspx";
                }
            }
            else
            {
                pageName = parser?.ParseString((clientSidePage as TranslatedClientSidePage)!.PageName) ?? (clientSidePage as TranslatedClientSidePage)!.PageName;
            }

            return pageName;
        }
#nullable restore


        // private (string url, Microsoft.SharePoint.Client.File file) PreCreatePage(Web web, ProvisioningTemplate template, TokenParser parser, BaseClientSidePage clientSidePage, string pagesLibrary, ref int currentPageIndex)
        // {
        //     string pageName = DeterminePageName(parser, clientSidePage);
        //     string url = $"{pagesLibrary}/{pageName}";

        //     if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
        //     {
        //         url = $"{pagesLibrary}/{dummyPage.GetTemplatesFolder()}/{pageName}";
        //     }

        //     // Write page level status messages, needed in case many pages are provisioned
        //     currentPageIndex++;
        //     WriteSubProgress("ClientSidePage", $"Create {pageName} stub", currentPageIndex, template.ClientSidePages.Count);

        //     url = UrlUtility.Combine(web.ServerRelativeUrl, url);

        //     var exists = true;
        //     try
        //     {
        //         var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
        //         web.Context.Load(file, f => f.UniqueId, f => f.ServerRelativePath, f => f.Exists);
        //         web.Context.ExecuteQueryRetry();

        //         // Fill token
        //         parser.AddToken(new PageUniqueIdToken(web, file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));
        //         parser.AddToken(new PageUniqueIdEncodedToken(web, file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));

        //         exists = file.Exists;
        //     }
        //     catch (ServerException ex)
        //     {
        //         if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
        //         {
        //             exists = false;
        //         }
        //     }

        //     if (!exists)
        //     {
        //         // Pre-create the page    
        //         PnPCore.IPage page = web.AddClientSidePage(pageName);

        //         // Set page layout now, because once it's set, it can't be changed.
        //         if (!string.IsNullOrEmpty(clientSidePage.Layout))
        //         {
        //             page.LayoutType = (PnPCore.PageLayoutType)Enum.Parse(typeof(PnPCore.PageLayoutType), clientSidePage.Layout);
        //         }

        //         string createdPageName;
        //         if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
        //         {
        //             createdPageName = page.SaveAsTemplate(pageName);
        //         }
        //         else
        //         {
        //             createdPageName = page.Save(pageName, HEUassumeListItemMissing: true);
        //         }

        //         url = $"{pagesLibrary}/{createdPageName}";
        //         if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
        //         {
        //             url = $"{pagesLibrary}/{dummyPage.GetTemplatesFolder()}/{pageName}";
        //         }
        //         url = UrlUtility.Combine(web.ServerRelativeUrl, url);

        //         var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
        //         web.Context.Load(file, f => f.UniqueId, f => f.ServerRelativePath);
        //         web.Context.ExecuteQueryRetry();

        //         // Fill token
        //         parser.AddToken(new PageUniqueIdToken(web, file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));
        //         parser.AddToken(new PageUniqueIdEncodedToken(web, file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));

        //         // Track that we pre-added this page
        //         return (url, file);
        //     }

        //     return (null, null);
        // }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (new PnPMonitoredScope(this.Name))
            {
                // Impossible to return all files in the site currently

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.ClientSidePages.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = false;
            }
            return _willExtract.Value;
        }
    }

#nullable enable
    public static class PagePreCreator
    {
        static readonly object _lock = new();
        // note: do NOT cache the actual file because pre-creation can be called from another thread and parallel to the actual page creation
        static readonly Dictionary<(string siteUrl, string finalPageName), Task<(string? pageUrl, List<Action<TokenParser, Web>> parserActions, bool? didExist)>> preCreatedFileUrls = new();

        // note that "didExists" can mean anything of the following: an actual WikiTraccs page with this name existed; a previously created, unfinished, precreated page exists; pre-page creation was scheduled by a previous call (result kind of unclear!?)
        public static Task<(string? pageUrl, List<Action<TokenParser, Web>> parserActions, bool? didExist)> PreCreateStandardWikiTraccsPageAsync(
            Web web,
            string pageName,
            string pagesLibrary,
            Action<string, string>? writeSubProgress)
        {
            var csp = new ClientSidePage()
            {
                PageName = pageName,
                PromoteAsTemplate = false,
            };
            return PreCreatePageAsync(web, csp, pagesLibrary, null, writeSubProgress);
        }

        public static Task<(string? pageUrl, List<Action<TokenParser, Web>> parserActions, bool? didExist)> PreCreatePageAsync(
            Web web,
            BaseClientSidePage clientSidePage,
            string pagesLibrary,
            Func<string>? getTemplatesFolderName,
            Action<string, string>? writeSubProgress)
        {
            var serverRelativeWebUrl = $"/{new Uri(web.Context.Url).AbsolutePath.TrimStart('/')}";
            var pageName = ObjectClientSidePages.DeterminePageName(null, clientSidePage);
            lock (_lock)
            {
                if (!preCreatedFileUrls.TryGetValue((serverRelativeWebUrl, pageName), out var task))
                {
                    task = PreCreatePageAsyncImpl(web, clientSidePage, pagesLibrary, getTemplatesFolderName, writeSubProgress);
                    preCreatedFileUrls[(serverRelativeWebUrl, pageName)] = task;
                }
                else
                {
                    task = Task.Run(async () =>
                    {
                        var (serverRelativePageUrl, parserActions, didExist) = await PreCreatePageAsyncImpl(web, clientSidePage, pagesLibrary, getTemplatesFolderName, writeSubProgress).ConfigureAwait(false);
                        // mark as "didExist"
                        return (serverRelativePageUrl, parserActions, (bool?)true);
                    });
                }
                return task;
            }
        }

        private static async Task<(string? serverRelativePageUrl, List<Action<TokenParser, Web>> parserActions, bool? didExist)> PreCreatePageAsyncImpl(
            Web web,
            BaseClientSidePage clientSidePage,
            string pagesLibrary,
            Func<string>? getTemplatesFolderName,
            Action<string, string>? writeSubProgress)
        {
            var parserActions = new List<Action<TokenParser, Web>>();
            if (clientSidePage.PromoteAsTemplate && null == getTemplatesFolderName)
            {
                // not supported
                return (null, parserActions, null);
            }
            string pageName = ObjectClientSidePages.DeterminePageName(null, clientSidePage);
            if (string.IsNullOrWhiteSpace(pageName))
            {
                throw new NotSupportedException("Need to set name of page to preload");
            }
            var serverRelativeWebUrl = $"/{new Uri(web.Context.Url).AbsolutePath.TrimStart('/')}";

            string url = $"{pagesLibrary}/{pageName}";

            if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
            {
                url = $"{pagesLibrary}/{getTemplatesFolderName!()}/{pageName}";
            }

            // Write page level status messages, needed in case many pages are provisioned

            //currentPageIndex++;
            writeSubProgress?.Invoke("ClientSidePage", $"Create {pageName} stub"); //, currentPageIndex, template.ClientSidePages.Count);


            url = UrlUtility.Combine(serverRelativeWebUrl, url);

            var exists = true;
            try
            {
                var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
                web.Context.Load(file, f => f.UniqueId, f => f.ServerRelativePath, f => f.Exists);
                await web.Context.ExecuteQueryRetryAsync().ConfigureAwait(false);

                var val1 = file.ServerRelativePath.DecodedUrl.Substring(serverRelativeWebUrl.Length).TrimStart("/".ToCharArray());
                var uniqueId1 = file.UniqueId;
                parserActions.Add((parser, otherWeb) => parser.AddToken(new PageUniqueIdToken(otherWeb, val1, uniqueId1)));
                parserActions.Add((parser, otherWeb) => parser.AddToken(new PageUniqueIdEncodedToken(otherWeb, val1, uniqueId1)));
                // Fill token
                // parser.AddToken(new PageUniqueIdToken(       web, file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));
                // parser.AddToken(new PageUniqueIdEncodedToken(web, file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));

                exists = file.Exists;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    exists = false;
                }
            }

            if (!exists)
            {
                // Pre-create the page    
                PnPCore.IPage page = await web.AddClientSidePageAsync(clientSidePage.EditorType, pageName).ConfigureAwait(false);

                // Set page layout now, because once it's set, it can't be changed.
                if (!string.IsNullOrEmpty(clientSidePage.Layout))
                {
                    page.LayoutType = (PnPCore.PageLayoutType)Enum.Parse(typeof(PnPCore.PageLayoutType), clientSidePage.Layout);
                }

                string createdPageName;
                if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
                {
                    createdPageName = page.SaveAsTemplate(pageName);
                }
                else
                {
                    createdPageName = page.Save(pageName, HEUassumeListItemMissing: true);
                }

                url = $"{pagesLibrary}/{createdPageName}";
                if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
                {
                    url = $"{pagesLibrary}/{getTemplatesFolderName!()}/{pageName}";
                }
                url = UrlUtility.Combine(serverRelativeWebUrl, url);

                var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
                web.Context.Load(file, f => f.UniqueId, f => f.ServerRelativePath);
                await web.Context.ExecuteQueryRetryAsync().ConfigureAwait(false);

                // Fill token
                // parser.AddToken(new PageUniqueIdToken(       web, file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));
                // parser.AddToken(new PageUniqueIdEncodedToken(web, file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));
                var val1 = file.ServerRelativePath.DecodedUrl.Substring(serverRelativeWebUrl.Length).TrimStart("/".ToCharArray());
                var uniqueId1 = file.UniqueId;
                parserActions.Add((parser, otherWeb) => parser.AddToken(new PageUniqueIdToken(otherWeb, val1, uniqueId1)));
                parserActions.Add((parser, otherWeb) => parser.AddToken(new PageUniqueIdEncodedToken(otherWeb, val1, uniqueId1)));

                return (url, parserActions, false);
            }

            return (url, parserActions, true);
        }        
    }    
#nullable restore
}
