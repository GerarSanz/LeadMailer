using LeadMailer.Helpers;
using LeadMailer.Models;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using System.IO;
using System.Net;

namespace LeadMailer.Services;

/// <summary>Envía correos HTML personalizados por curso.</summary>
public class EmailService
{
    public (string Subject, string HtmlBody, string TextBody) BuildPreview(Lead lead, CourseInfo course, SmtpConfig cfg)
    {
        var nombreCursoEmail = CourseNameForEmail(course);
        var subject = string.IsNullOrWhiteSpace(course.AsuntoEmail)
            ? $"Información sobre: {nombreCursoEmail}"
            : course.AsuntoEmail;

        return (
            Subject: subject,
            HtmlBody: BuildHtml(lead, course, cfg),
            TextBody: BuildText(lead, course, cfg));
    }

    // ── Envío individual ──────────────────────────────────────────────────────
    public async Task<(bool Success, string? Error)> SendAsync(Lead lead, CourseInfo course, SmtpConfig cfg)
    {
        try
        {
            var msg = new MimeMessage();
            msg.From.Add(new MailboxAddress(cfg.FromName, cfg.FromEmail));
            msg.To.Add(new MailboxAddress(lead.Nombre, lead.Email));
            var nombreCursoEmail = CourseNameForEmail(course);
            msg.Subject = string.IsNullOrWhiteSpace(course.AsuntoEmail)
                ? $"Información sobre: {nombreCursoEmail}"
                : course.AsuntoEmail;

            var builder = new BodyBuilder
            {
                HtmlBody  = BuildHtml(lead, course, cfg),
                TextBody  = BuildText(lead, course, cfg)
            };

            if (!string.IsNullOrWhiteSpace(course.PdfAdjuntoPath))
            {
                var pdfPath = course.PdfAdjuntoPath.Trim();
                if (!File.Exists(pdfPath))
                    return (false, $"No se encontró el PDF adjunto configurado para el curso: {pdfPath}");

                builder.Attachments.Add(pdfPath);
            }

            msg.Body = builder.ToMessageBody();

            using var client = new SmtpClient();
            var ssl = cfg.EnableSsl ? SecureSocketOptions.StartTls : SecureSocketOptions.None;
            await client.ConnectAsync(cfg.Host, cfg.Port, ssl);
            await client.AuthenticateAsync(cfg.Username, cfg.Password);
            await client.SendAsync(msg);
            await client.DisconnectAsync(true);

            return (true, null);
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    public async Task<(bool Success, string? Error)> SendReportAsync(
        SmtpConfig cfg,
        string toEmail,
        string subject,
        string bodyText,
        string attachmentFileName,
        byte[] attachmentBytes)
    {
        try
        {
            var msg = new MimeMessage();
            var fromEmail = string.IsNullOrWhiteSpace(cfg.FromEmail) ? cfg.Username : cfg.FromEmail;
            var fromName = string.IsNullOrWhiteSpace(cfg.FromName) ? "LeadMailer" : cfg.FromName;

            msg.From.Add(new MailboxAddress(fromName, fromEmail));
            msg.To.Add(MailboxAddress.Parse(toEmail));
            msg.Subject = subject;

            var builder = new BodyBuilder { TextBody = bodyText };
            builder.Attachments.Add(
                attachmentFileName,
                attachmentBytes,
                ContentType.Parse("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

            msg.Body = builder.ToMessageBody();

            using var client = new SmtpClient();
            var ssl = cfg.EnableSsl ? SecureSocketOptions.StartTls : SecureSocketOptions.None;
            await client.ConnectAsync(cfg.Host, cfg.Port, ssl);
            await client.AuthenticateAsync(cfg.Username, cfg.Password);
            await client.SendAsync(msg);
            await client.DisconnectAsync(true);

            return (true, null);
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    // ── Test de conexión ──────────────────────────────────────────────────────
    public async Task<(bool Success, string? Error)> TestConnectionAsync(SmtpConfig cfg)
    {
        try
        {
            using var client = new SmtpClient();
            var ssl = cfg.EnableSsl ? SecureSocketOptions.StartTls : SecureSocketOptions.None;
            await client.ConnectAsync(cfg.Host, cfg.Port, ssl);
            await client.AuthenticateAsync(cfg.Username, cfg.Password);
            await client.DisconnectAsync(true);
            return (true, null);
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    // ── Construcción del HTML ─────────────────────────────────────────────────
    private static string BuildHtml(Lead lead, CourseInfo c, SmtpConfig cfg)
    {
        var nombre   = string.IsNullOrWhiteSpace(lead.Nombre) ? "Estimado/a" : lead.Nombre;
        var nombreCursoEmail = CourseNameForEmail(c);

        var richMarketingHtml = !string.IsNullOrWhiteSpace(c.TextoMarketingRich)
            ? RichTextSerialization.XamlToHtml(c.TextoMarketingRich)
            : string.Empty;

        var mktHtml  = !string.IsNullOrWhiteSpace(richMarketingHtml)
            ? $"<div style='font-size:15px;line-height:1.75;color:#374151'>{richMarketingHtml}</div>"
            : (!string.IsNullOrWhiteSpace(c.TextoMarketing)
                ? $"<p style='font-size:15px;line-height:1.75;color:#374151'>{Enc(c.TextoMarketing).Replace("\n","<br/>")}</p>"
                : "");

        var filas = new System.Text.StringBuilder();
        void Fila(string icon, string label, string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
                filas.Append($@"<tr>
                  <td style='padding:10px 14px;border-bottom:1px solid #E5E7EB;color:#6B7280;
                             font-size:13px;white-space:nowrap;vertical-align:top;font-weight:600'>
                    {icon}&nbsp;{label}
                  </td>
                  <td style='padding:10px 14px;border-bottom:1px solid #E5E7EB;font-size:13px;
                             color:#111827;vertical-align:top'>
                    {Enc(value).Replace("\n","<br/>")}
                  </td>
                </tr>");
        }

        Fila("📅", "Fecha de inicio",        c.FechaInicio);
        Fila("📅", "Fecha de fin",            c.FechaFin);
        Fila("🕐", "Horario",                 c.HorarioInfo);
        Fila("✅", "Requisitos de acceso",    c.RequisitosAcceso);
        Fila("📄", "Documentación necesaria", c.DocumentacionNecesaria);
        Fila("📎", "Diseño formativo oficial", c.UrlFichaInscripcion);
        Fila("ℹ️", "Información adicional",   c.InfoAdicional);

        var tabla = filas.Length > 0 ? $@"
          <h3 style='font-size:15px;font-weight:700;color:#1E40AF;margin:24px 0 12px'>Detalles del curso</h3>
          <table style='width:100%;border-collapse:collapse;background:#F9FAFB;border-radius:8px;overflow:hidden'>
            {filas}
          </table>" : "";

        // Botón WhatsApp si está configurado
        var whatsappButton = string.Empty;
        if (!string.IsNullOrWhiteSpace(cfg.WhatsAppNumber))
        {
            var whatsappPhone = BuildWhatsAppPhone(cfg.WhatsAppNumber);
            if (!string.IsNullOrWhiteSpace(whatsappPhone))
            {
                whatsappButton = $@"<p style='margin:12px 0 0;'>
                  <a href='https://wa.me/{whatsappPhone}'
                     style='display:inline-block;background:#25D366;color:white;padding:12px 24px;
                            border-radius:8px;text-decoration:none;font-weight:600;font-size:14px'>
                    💬 Chatea por WhatsApp
                  </a>
                </p>";
            }
        }

        var footerHtml = !string.IsNullOrWhiteSpace(cfg.EmailFooter)
            ? $@"<div style='margin-top:14px;font-size:11px;color:#6B7280;line-height:1.6'>
                 {Enc(cfg.EmailFooter).Replace("\n", "<br/>")}
               </div>"
            : "";

        return $@"<!DOCTYPE html>
<html lang='es'><head><meta charset='UTF-8'></head>
<body style='margin:0;padding:0;background:#F3F4F6;font-family:Segoe UI,Arial,sans-serif'>
  <table width='100%' cellpadding='0' cellspacing='0'>
    <tr><td align='center' style='padding:32px 16px'>
      <table width='620' cellpadding='0' cellspacing='0'
             style='background:white;border-radius:16px;overflow:hidden;
                    box-shadow:0 4px 24px rgba(0,0,0,0.08)'>
        <tr><td style='background:linear-gradient(135deg,#1E40AF,#0EA5E9);padding:20px 36px;text-align:center'>
          <img src='https://grupoaspasia.com/vue/_nuxt/logo_azul.CkcnliBr.svg' alt='Grupo Aspasia' style='height:40px;margin-bottom:12px'>
          <p style='color:rgba(255,255,255,0.7);margin:0;font-size:11px;
                    text-transform:uppercase;letter-spacing:1.5px'>Formación profesional</p>
          <h1 style='color:white;margin:8px 0 0;font-size:22px;font-weight:700;line-height:1.3'>
            {Enc(nombreCursoEmail)}
          </h1>
        </td></tr>
        <tr><td style='padding:32px 36px'>
          <p style='font-size:16px;color:#111827;margin:0 0 6px'>
            Hola <strong>{Enc(nombre)}</strong>,
          </p>
          <p style='font-size:14px;color:#6B7280;margin:0 0 24px'>
            Gracias por tu interés. A continuación te enviamos toda la información sobre la formación.
          </p>
          {mktHtml}
          {tabla}
        </td></tr>
        <tr><td style='background:#F9FAFB;border-top:1px solid #E5E7EB;
                       padding:24px 36px;text-align:center'>
          <p style='margin:0 0 12px;font-size:13px;color:#6B7280;font-weight:600'>
            ¿Tienes preguntas? Contacta directamente con nosotros
          </p>
          {whatsappButton}
          {footerHtml}
        </td></tr>
      </table>
    </td></tr>
  </table>
</body></html>";
    }

    private static string BuildWhatsAppPhone(string? phone)
    {
        if (string.IsNullOrWhiteSpace(phone)) return "";
        
        // Extraer solo dígitos y eliminar espacios/caracteres especiales
        var digits = new string(phone.Where(char.IsDigit).ToArray());
        if (string.IsNullOrWhiteSpace(digits)) return "";

        // Eliminar prefijos + o 00
        if (digits.StartsWith("00"))
            digits = digits.Substring(2);

        // Si comienza con 34 (código de España)
        if (digits.StartsWith("34"))
        {
            // Aceptar: 34 + 8-9 dígitos = 10-11 dígitos total
            if (digits.Length >= 10 && digits.Length <= 11)
                return digits;
            // Rechazar si tiene demasiados dígitos (posible duplicado)
            if (digits.Length > 11)
                return "";
            // Rechazar si tiene muy pocos dígitos
            if (digits.Length < 10)
                return "";
        }

        // Caso España: número de 9 dígitos sin prefijo (móvil: 6xx, 7xx o fijo: 9xx)
        if (digits.Length == 9)
        {
            var firstDigit = digits[0];
            // Validar que comience con dígito válido español (6, 7, 8, 9)
            if ("6789".Contains(firstDigit.ToString()))
                return $"34{digits}";
            else
                return ""; // Número inválido
        }

        // Otros formatos internacionales válidos (10-15 dígitos)
        if (digits.Length >= 10 && digits.Length <= 15)
            return digits;

        // Números inválidos (menos de 9 dígitos)
        return "";
    }

    // ── Versión texto plano ───────────────────────────────────────────────────
    private static string BuildText(Lead lead, CourseInfo c, SmtpConfig cfg)
    {
        var sb     = new System.Text.StringBuilder();
        var nombre = string.IsNullOrWhiteSpace(lead.Nombre) ? "Estimado/a" : lead.Nombre;
        var nombreCursoEmail = CourseNameForEmail(c);
        sb.AppendLine($"Hola {nombre},");
        sb.AppendLine();
        sb.AppendLine($"Gracias por tu interés en: {nombreCursoEmail}");
        sb.AppendLine();
        var marketingText = !string.IsNullOrWhiteSpace(c.TextoMarketing)
            ? c.TextoMarketing
            : RichTextSerialization.XamlToPlainText(c.TextoMarketingRich);
        if (!string.IsNullOrWhiteSpace(marketingText)) { sb.AppendLine(marketingText); sb.AppendLine(); }
        if (!string.IsNullOrWhiteSpace(c.FechaInicio))        sb.AppendLine($"Inicio:          {c.FechaInicio}");
        if (!string.IsNullOrWhiteSpace(c.FechaFin))           sb.AppendLine($"Fin:             {c.FechaFin}");
        if (!string.IsNullOrWhiteSpace(c.HorarioInfo))        sb.AppendLine($"Horario:         {c.HorarioInfo}");
        if (!string.IsNullOrWhiteSpace(c.RequisitosAcceso))   sb.AppendLine($"Requisitos:      {c.RequisitosAcceso}");
        if (!string.IsNullOrWhiteSpace(c.DocumentacionNecesaria)) sb.AppendLine($"Documentación:   {c.DocumentacionNecesaria}");
        if (!string.IsNullOrWhiteSpace(c.UrlFichaInscripcion)) sb.AppendLine($"Inscripción:     {c.UrlFichaInscripcion}");
        if (!string.IsNullOrWhiteSpace(c.InfoAdicional))      sb.AppendLine($"Más información: {c.InfoAdicional}");
        if (!string.IsNullOrWhiteSpace(cfg.EmailFooter))
        {
            sb.AppendLine();
            sb.AppendLine(cfg.EmailFooter.Trim());
        }
        return sb.ToString();
    }

    private static string CourseNameForEmail(CourseInfo c)
        => string.IsNullOrWhiteSpace(c.NombreComplementario) ? c.NombreCorto : c.NombreComplementario;

    private static string Enc(string s) => WebUtility.HtmlEncode(s ?? "");
}
