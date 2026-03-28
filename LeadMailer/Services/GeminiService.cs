using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;

namespace LeadMailer.Services;

public enum AiProvider { OpenRouter, Gemini }

/// <summary>
/// Genera texto de marketing.
/// - OpenRouter (GRATUITO): meta-llama/llama-3.1-8b-instruct:free  → openrouter.ai
/// - Gemini (Google AI Studio): gemini-3-flash-preview             → aistudio.google.com
/// </summary>
public class AiTextService
{
    // v1beta es necesario para los modelos 2.0+; también soporta los 1.5
    private const string GeminiBaseUrl =
        "https://generativelanguage.googleapis.com/v1beta/models/";

    // Modelos Gemini en orden de preferencia (de más nuevo a más antiguo)
    private static readonly string[] GeminiModels =
    [
        "gemini-3-flash-preview",
        "gemini-2.0-flash",
        "gemini-2.0-flash-lite",
        "gemini-1.5-flash",
        "gemini-1.5-flash-8b",
    ];

    private const string OpenRouterUrl = "https://openrouter.ai/api/v1/chat/completions";

    // Modelos gratuitos de OpenRouter en orden de preferencia.
    // Si el primero no tiene endpoints disponibles, se prueba el siguiente.
    private static readonly string[] OpenRouterFreeModels =
    [
        "meta-llama/llama-3.1-8b-instruct:free",
        "meta-llama/llama-3.2-3b-instruct:free",
        "mistralai/mistral-7b-instruct:free",
        "google/gemma-2-9b-it:free",
    ];

    private static readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(45) };

    public async Task<(bool Success, string Result)> GenerateMarketingTextAsync(
        string     apiKey,
        AiProvider provider,
        string     nombreCurso,
        string     requisitos,
        string     fechaInicio,
        string     fechaFin,
        string     horario,
        string     infoAdicional)
    {
        if (string.IsNullOrWhiteSpace(apiKey))
            return (false, "⚠  Introduce la API Key en la pestaña Configuración.");

        var prompt = BuildPrompt(nombreCurso, requisitos, fechaInicio, fechaFin, horario, infoAdicional);

        return provider == AiProvider.Gemini
            ? await CallGeminiAsync(apiKey, prompt)
            : await CallOpenRouterAsync(apiKey, prompt);
    }

    // ── OpenRouter ────────────────────────────────────────────────────────────
    private static async Task<(bool, string)> CallOpenRouterAsync(string apiKey, string prompt)
    {
        string? lastError = null;

        foreach (var model in OpenRouterFreeModels)
        {
            try
            {
                var body = new
                {
                    model    = model,
                    messages = new[]
                    {
                        new { role = "system", content = "Eres un experto en marketing de formación profesional en España. Respondes ÚNICAMENTE con el texto de marketing solicitado, sin explicaciones, sin saludos, sin formato markdown ni asteriscos." },
                        new { role = "user",   content = prompt }
                    },
                    max_tokens  = 600,
                    temperature = 0.8
                };

                var req = new HttpRequestMessage(HttpMethod.Post, OpenRouterUrl);
                req.Headers.Add("Authorization", $"Bearer {apiKey}");
                req.Headers.Add("HTTP-Referer", "https://leadmailer.app");
                req.Headers.Add("X-Title", "LeadMailer");
                req.Content = new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json");

                var response = await _http.SendAsync(req);
                var raw      = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    var msg = TryGetErrorMessage(raw) ?? raw;
                    // "No endpoints" significa que el modelo no está disponible ahora → prueba el siguiente
                    if (msg.Contains("No endpoints", StringComparison.OrdinalIgnoreCase))
                    {
                        lastError = $"Modelo '{model}' no disponible.";
                        continue;
                    }
                    return (false, $"Error OpenRouter ({model}): {msg}");
                }

                var text = JObject.Parse(raw)["choices"]?[0]?["message"]?["content"]?.ToString();
                if (string.IsNullOrWhiteSpace(text))
                {
                    lastError = $"Modelo '{model}' no devolvió texto.";
                    continue;
                }

                return (true, text.Trim());
            }
            catch (HttpRequestException ex) when (ex.InnerException is SocketException se && se.SocketErrorCode == SocketError.HostNotFound)
            {
                return (false, "No se pudo resolver el servidor de OpenRouter. Revisa tu conexión a Internet.");
            }
            catch (TaskCanceledException)
            {
                return (false, "Tiempo de espera agotado. Comprueba tu conexión.");
            }
            catch (Exception ex)
            {
                lastError = $"Error en '{model}': {ex.Message}";
            }
        }

        return (false, lastError ?? "Ningún modelo gratuito de OpenRouter está disponible ahora. Prueba más tarde o usa Gemini.");
    }

    // ── Gemini ───────────────────────────────────────────────────────────────
    private static async Task<(bool, string)> CallGeminiAsync(string apiKey, string prompt)
    {
        string? lastError = null;

        foreach (var model in GeminiModels)
        {
            try
            {
                var body = new
                {
                    contents         = new[] { new { parts = new[] { new { text = prompt } } } },
                    generationConfig = new { temperature = 0.8, maxOutputTokens = 600, topP = 0.9 }
                };

                var url      = $"{GeminiBaseUrl}{model}:generateContent?key={apiKey}";
                var content  = new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json");
                var response = await _http.PostAsync(url, content);
                var raw      = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    var msg = TryGetErrorMessage(raw) ?? raw;
                    // Modelo no encontrado o no disponible → prueba el siguiente
                    if (response.StatusCode == System.Net.HttpStatusCode.NotFound ||
                        msg.Contains("not found", StringComparison.OrdinalIgnoreCase) ||
                        msg.Contains("deprecated", StringComparison.OrdinalIgnoreCase))
                    {
                        lastError = $"Modelo '{model}' no disponible: {msg}";
                        continue;
                    }
                    return (false, $"Error Gemini ({model}): {msg}");
                }

                var text = JObject.Parse(raw)["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.ToString();
                if (string.IsNullOrWhiteSpace(text))
                {
                    lastError = $"Modelo '{model}' no devolvió texto.";
                    continue;
                }

                return (true, text.Trim());
            }
            catch (TaskCanceledException) { return (false, "Tiempo de espera agotado. Comprueba tu conexión."); }
            catch (Exception ex)          { lastError = $"Error en '{model}': {ex.Message}"; }
        }

        return (false, lastError ?? "Ningún modelo de Gemini está disponible. Comprueba tu API Key o prueba más tarde.");
    }

    // ── Helpers ───────────────────────────────────────────────────────────────
    private static string? TryGetErrorMessage(string raw)
    {
        try { return JObject.Parse(raw)["error"]?["message"]?.ToString(); }
        catch { return null; }
    }

    private static string BuildPrompt(string nombre, string requisitos,
        string inicio, string fin, string horario, string extra)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Escribe un texto de marketing atractivo para un correo a personas interesadas en este curso de formación profesional en España.");
        sb.AppendLine("REGLAS: tono cercano y motivador, 80-150 palabras, destaca beneficios, sin saludos ni despedidas, español de España, sin asteriscos ni markdown.");
        sb.AppendLine();
        sb.AppendLine($"Curso: {nombre}");
        if (!string.IsNullOrWhiteSpace(requisitos)) sb.AppendLine($"Requisitos: {requisitos}");
        if (!string.IsNullOrWhiteSpace(inicio))     sb.AppendLine($"Inicio: {inicio}");
        if (!string.IsNullOrWhiteSpace(fin))        sb.AppendLine($"Fin: {fin}");
        if (!string.IsNullOrWhiteSpace(horario))    sb.AppendLine($"Horario: {horario}");
        if (!string.IsNullOrWhiteSpace(extra))      sb.AppendLine($"Info: {extra}");
        return sb.ToString();
    }
}
