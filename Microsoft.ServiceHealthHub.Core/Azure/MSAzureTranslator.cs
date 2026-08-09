using Microsoft.ServiceHealthHub.Core;
using Newtonsoft.Json;
using System.Text;

namespace Microsoft.ServiceHealthHub.Azure
{
    public class MSAzureTranslation
    {
        public string text { get; set; }
        public string to { get; set; }
    }

    public class MSAzureTranslationDetectedLanguage
    {
        public string language { get; set; }
        public string score { get; set; }
    }

    public class MSAzureTranslationResult
    {
        public List<MSAzureTranslation> translations { get; set; }
    }

    public class MSAzureTranslator
    {
        private static readonly string subscriptionKey = GlobalConfiguration.Instance.AzureTranslatorSubscriptionKey;
        private static readonly string endpoint = "https://api.cognitive.microsofttranslator.com/";
        private static readonly string location = GlobalConfiguration.Instance.AzureTranslatorResourceLocation;
        private static readonly MSSHNotificationDatabase db = new MSSHNotificationDatabase();

        private static async Task<List<MSTranslationCollection>> GetTranslationFromAPI(string[] messages, string[] languages, bool html=false)
        {
            object[] body;
            List<object> listOfMessages = new List<object>();
            List<MSTranslationCollection> result = new List<MSTranslationCollection>();

            string route = "/translate?api-version=3.0";

            foreach (string language in languages)
            {
                route += string.Format("&to={0}", language);
            }

            if (html)
                route += "&textType=html";

            foreach (string message in messages)
            {
                listOfMessages.Add(new { Text = message });
            }
            body = listOfMessages.ToArray();
            var requestBody = JsonConvert.SerializeObject(body);

            using (var client = new HttpClient())
            using (var request = new HttpRequestMessage())
            {
                // Build the request.
                request.Method = HttpMethod.Post;
                request.RequestUri = new Uri(endpoint + route);
                request.Content = new StringContent(requestBody, Encoding.UTF8, "application/json");
                request.Headers.Add("Ocp-Apim-Subscription-Key", subscriptionKey);
                request.Headers.Add("Ocp-Apim-Subscription-Region", location);

                // Send the request and get response.
                HttpResponseMessage response = await client.SendAsync(request).ConfigureAwait(false);
                // Read response as a string.
                string httpResult = await response.Content.ReadAsStringAsync();

                List<MSAzureTranslationResult> deserializedResult = null;
                try
                {
                    deserializedResult = JsonConvert.DeserializeObject<List<MSAzureTranslationResult>>(httpResult);
                } catch(Exception ex)
                {
                    Console.WriteLine("");
                }

                int msgCount = 0;
                foreach (MSAzureTranslationResult translationResult in deserializedResult)
                {
                    MSTranslationCollection item = new MSTranslationCollection();
                    item.OriginalMessage = messages[msgCount];

                    foreach (var translation in translationResult.translations)
                    {
                        MSTranslation translationItem = new MSTranslation();
                        translationItem.Language = translation.to;
                        translationItem.Message = translation.text;
                        item.Translations.Add(translationItem.Language, translationItem);
                    }
                    result.Add(item);

                    msgCount++;
                }
            }

            return result;
        }

        public static List<MSTranslationCollection> Translate(string[] messages, string language, bool html=false)
        {
            List<string> notFoundInCache = new List<string>();
            List<MSTranslationCollection> result = new List<MSTranslationCollection>();

            List<MSTranslationCollection> translationCache = db.GetTranslationFromCache(messages, language);

            foreach (string message in messages)
            {
                MSTranslationCollection translation = translationCache.Find(t => t.OriginalMessage == message);

                if (translation == null)
                {
                    notFoundInCache.Add(message);
                }
                else
                {
                    result.Add(translation);
                }
            }

            if (notFoundInCache.Count > 0)
            {
                Task<List<MSTranslationCollection>> t = Task.Run(async () =>
                {
                    List<MSTranslationCollection> res = await MSAzureTranslator.GetTranslationFromAPI(notFoundInCache.ToArray(), new string[] { language }, html);
                    return res;
                });

                t.Wait();

                List<MSTranslationCollection> apiResult = t.Result;

                foreach (MSTranslationCollection translationCollection in apiResult)
                {
                    result.Add(translationCollection);
                    foreach (string lang in translationCollection.Translations.Keys)
                    {
                        db.CacheTranslation(translationCollection.OriginalMessage, translationCollection.Translations[lang]);
                    }
                }
            }

            return result;
        }
    }
}
