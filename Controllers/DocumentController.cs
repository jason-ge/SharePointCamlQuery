using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SharePointCamlQuery.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace SharePointCamlQuery.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class DocumentController : ControllerBase
    {
        private readonly ILogger<DocumentController> _logger;
        private readonly IConfiguration _configuration;
        private readonly HttpClient _httpClient;
        private const string CAML_QUERY = "{{ 'query' :{{" + 
            "'__metadata': {{ 'type': 'SP.CamlQuery' }}, " +
            "'ViewXml': " + 
                "'<View Scope=\"Recursive\">" +
                    "<ViewFields>" +
                        "<FieldRef Name=\"Created\"></FieldRef><FieldRef Name=\"FileLeafRef\"></FieldRef>" +
                        "<FieldRef Name=\"Modified\"></FieldRef>" +
                        "<FieldRef Name=\"ID\"></FieldRef>" +
                        "<FieldRef Name=\"Document_x0020_Type\"></FieldRef>" +
                        "<FieldRef Name=\"EncodedAbsUrl\"></FieldRef>" +
                    "</ViewFields>" +
                    "<Query>" +
                        "{0}" +
                    "</Query>" +
                "</View>'" +
            "}}}}";
        private const string DOCUMENT_SEARCH_API_ENDPOINT = "{0}/_api/web/lists/getbytitle('{1}')/GetItems?$select=FileLeafRef,Created,Modified,ID,EncodedAbsUrl";
        private const string METADATA_API_ENDPOINT = "{0}/_api/web/lists/getbytitle('{1}')/Fields?$select=Title,InternalName";

        public DocumentController(IConfiguration configuration, ILogger<DocumentController> logger, IHttpClientFactory factory)
        {
            _httpClient = factory.CreateClient("SharepointClient");
            _logger = logger;
            _configuration = configuration;
        }

        /// <summary>
        /// Gets the Document Libraries from SharePoint Site, 
        /// </summary>
        /// <returns>List of Document Libraries</returns>
        [HttpGet("getdocumentlibraries")]
        public async Task<IEnumerable<ListInfo>> GetDocumentLibraryListAsync()
        {
            const string LIST_QUERY_API_ENDPOINT = "{0}/_api/web/lists?$filter=Hidden+eq+false%20and%20BaseType+eq+1%20and%20BaseTemplate+eq+101";
            var siteUrl = _configuration["SharePointSettings:SiteUrl"];
            var queryApiUrl = string.Format(LIST_QUERY_API_ENDPOINT, siteUrl);
            var response = await QuerySharepoint(siteUrl, queryApiUrl, HttpMethod.Get);
            if ((int)response.StatusCode == 200)
            {
                var results = JsonConvert.DeserializeObject<ListQueryResponse>(await response.Content.ReadAsStringAsync());
                return results.d.Results;
            }
            else
            {
                // Log error here
                _logger.LogError($"{nameof(GetDocumentLibraryListAsync)} error: {JsonConvert.SerializeObject(response)}");
                return new List<ListInfo>();
            }
        }

        [HttpGet("/{documentLibraryName}/search/{titlePartial}")]
        public async Task<IEnumerable<DocumentModel>> SearchDocumentLibraryAsync([FromRoute]string documentLibraryName, [FromRoute]string titlePartial)
        {
            var siteUrl = _configuration["SharePointSettings:SiteUrl"];
            var caml = string.Format(CAML_QUERY,
                            "<Where><Contains><FieldRef Name=\"FileLeafRef\" /><Value Type=\"Text\">" + titlePartial + "</Value></Contains></Where>");
            return await QueryDocumentLibray(siteUrl, string.Format(DOCUMENT_SEARCH_API_ENDPOINT, siteUrl, documentLibraryName), HttpMethod.Post, caml);
        }

        [HttpGet("/{documentLibraryName}/fields")]
        public async Task<IEnumerable<FieldInfo>> GetDocumentLibraryFielsAsync([FromRoute] string documentLibraryName)
        {
            var siteUrl = _configuration["SharePointSettings:SiteUrl"];
            var response = await QuerySharepoint(siteUrl, string.Format(METADATA_API_ENDPOINT, siteUrl, documentLibraryName), HttpMethod.Get);
            if ((int)response.StatusCode == 200)
            {
                var results = JsonConvert.DeserializeObject<FieldInfoResponse>(await response.Content.ReadAsStringAsync());
                return results.d.Results;
            }
            else
            {
                _logger.LogError($"{nameof(GetDocumentLibraryFielsAsync)} error.");
                return null;
            }
        }

        [HttpGet("/{documentLibraryName}/download/{documentId:int}")]
        public async Task<ActionResult> DownloadDocumentAsync([FromRoute] string documentLibraryName, [FromRoute] int documentId)
        {
            var siteUrl = _configuration["SharePointSettings:SiteUrl"];
            var caml = string.Format(CAML_QUERY,
                "<Where><Eq><FieldRef Name=\"ID\" /><Value Type=\"Integer\">" + documentId + "</Value></Eq></Where>");
            var documents = await QueryDocumentLibray(siteUrl, string.Format(DOCUMENT_SEARCH_API_ENDPOINT, siteUrl, documentLibraryName), HttpMethod.Post, caml);
            if (documents == null || documents.Count() == 0)
            {
                _logger.LogError($"{nameof(DownloadDocumentAsync)} cannot find the document with CAML query: {caml}");
                return null;
            }
            else if (documents.Count() > 1)
            {
                _logger.LogError($"{nameof(DownloadDocumentAsync)} finds {documents.Count()} documents with CAML query: {caml}");
                return null;
            }
            else
            {
                var document = documents.First();
                try
                {
                    return File(await _httpClient.GetByteArrayAsync(document.EncodedAbsUrl), "application/pdf");
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"{nameof(DownloadDocumentAsync)} error.");
                    return null;
                }
            }
        }

        private async Task<HttpResponseMessage> QuerySharepoint(string siteUrl, string docLibUrl, HttpMethod method, string caml = null)
        {
            var request = new HttpRequestMessage { Method = method, RequestUri = new Uri(docLibUrl) };
            request.Headers.Add("ContentType", "application/json;odata=verbose");
            request.Headers.Add("Accept", "application/json;odata=verbose");
            request.Headers.Add("X-RequestDigest", await GetFormDigestAsync(siteUrl).ConfigureAwait(false));
            if (caml != null)
            {
                request.Content = new StringContent(caml, Encoding.UTF8);
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            }

            return await _httpClient.SendAsync(request).ConfigureAwait(false);
        }

        private async Task<IEnumerable<DocumentModel>> QueryDocumentLibray(string siteUrl, string docLibUrl, HttpMethod method, string caml = null)
        {
            try
            {
                var response = await QuerySharepoint(siteUrl, docLibUrl, method, caml);
                if ((int)response.StatusCode == 200)
                {
                    var results = JsonConvert.DeserializeObject<DocumentQueryResponse>(await response.Content.ReadAsStringAsync());
                    return results.d.Results;
                }
                else
                {
                    // Log error here
                    return new List<DocumentModel>();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"{nameof(QueryDocumentLibray)} error: Url: {docLibUrl}, Query: {caml ?? ""}");
                return new List<DocumentModel>();
            }
        }

        private async Task<string> GetFormDigestAsync(string siteUrl)
        {
            const string FormDigestValue_Key = "FormDigestValue";
            var request = new HttpRequestMessage { Method = HttpMethod.Post, RequestUri = new Uri(siteUrl + "/_api/contextinfo") };
            request.Headers.Add("ContentType", "application/json");
            request.Headers.Add("Accept", "application/json");

            var response = await _httpClient.SendAsync(request).ConfigureAwait(false);
            if ((int)response.StatusCode == 200)
            {
                var result = JsonConvert.DeserializeObject<Dictionary<string, object>>(await response.Content.ReadAsStringAsync());
                return result != null && result.ContainsKey(FormDigestValue_Key) ? result[FormDigestValue_Key].ToString() : null;
            }
            else
            {
                return null;
            }
        }
    }
}
