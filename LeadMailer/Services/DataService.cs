using LeadMailer.Models;
using Newtonsoft.Json;
using System.IO;

namespace LeadMailer.Services;

/// <summary>Persistencia de cursos, configuración SMTP y registros de envío.</summary>
public class DataService
{
    private static readonly string DataFolder =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LeadMailer");
    private static readonly string DataFile = Path.Combine(DataFolder, "data.json");

    private AppData _data = new();

    public AppData     Data        => _data;
    public SmtpConfig  SmtpConfig  => _data.SmtpConfig;
    public List<CourseInfo>  Courses     => _data.Courses;
    public List<Student>     Students    => _data.Students;
    public List<SentRecord>  SentRecords => _data.SentRecords;

    public DataService() => Load();

    // ── CRUD ──────────────────────────────────────────────────────────────────

    public void Load()
    {
        try
        {
            Directory.CreateDirectory(DataFolder);
            if (File.Exists(DataFile))
            {
                var json = File.ReadAllText(DataFile);
                _data = NormalizeData(JsonConvert.DeserializeObject<AppData>(json) ?? new AppData());
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("DataService.Load: no se pudo cargar data.json, se usará configuración vacía", ex);
            _data = new AppData();
        }
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(DataFolder);
            File.WriteAllText(DataFile, JsonConvert.SerializeObject(_data, Formatting.Indented));
        }
        catch (Exception ex)
        {
            AppLogger.Error("DataService.Save: no se pudo guardar data.json", ex);
        }
    }

    public void ExportData(string destinationFilePath)
    {
        if (string.IsNullOrWhiteSpace(destinationFilePath))
            throw new ArgumentException("Ruta de exportación no válida.", nameof(destinationFilePath));

        Save();
        var folder = Path.GetDirectoryName(destinationFilePath);
        if (!string.IsNullOrWhiteSpace(folder))
            Directory.CreateDirectory(folder);

        File.Copy(DataFile, destinationFilePath, overwrite: true);
    }

    public void ImportData(string sourceFilePath)
    {
        if (string.IsNullOrWhiteSpace(sourceFilePath) || !File.Exists(sourceFilePath))
            throw new FileNotFoundException("No se encontró el archivo de importación.", sourceFilePath);

        var json = File.ReadAllText(sourceFilePath);
        var imported = JsonConvert.DeserializeObject<AppData>(json)
            ?? throw new InvalidDataException("El archivo no tiene un formato válido.");

        _data = NormalizeData(imported);
        Save();
    }

    public void ExportSmtpConfig(string destinationFilePath)
    {
        if (string.IsNullOrWhiteSpace(destinationFilePath))
            throw new ArgumentException("Ruta de exportación no válida.", nameof(destinationFilePath));

        var folder = Path.GetDirectoryName(destinationFilePath);
        if (!string.IsNullOrWhiteSpace(folder))
            Directory.CreateDirectory(folder);

        File.WriteAllText(destinationFilePath, JsonConvert.SerializeObject(_data.SmtpConfig, Formatting.Indented));
    }

    public void ImportSmtpConfig(string sourceFilePath)
    {
        if (string.IsNullOrWhiteSpace(sourceFilePath) || !File.Exists(sourceFilePath))
            throw new FileNotFoundException("No se encontró el archivo de configuración.", sourceFilePath);

        var json = File.ReadAllText(sourceFilePath);
        var imported = JsonConvert.DeserializeObject<SmtpConfig>(json)
            ?? throw new InvalidDataException("El archivo no tiene un formato de configuración válido.");

        _data.SmtpConfig = imported;
        Save();
    }

    // ── Helpers cursos ────────────────────────────────────────────────────────

    public CourseInfo? GetCourseByRaw(string cursoRaw)
        => _data.Courses.FirstOrDefault(c =>
            string.Equals(c.CursoRaw?.Trim(), cursoRaw?.Trim(), StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Devuelve el curso existente o crea uno nuevo si no existe.
    /// Si es nuevo, lo añade a la lista y guarda.
    /// </summary>
    public CourseInfo GetOrCreateCourse(string cursoRaw)
    {
        var existing = GetCourseByRaw(cursoRaw);
        if (existing != null) return existing;

        var course = new CourseInfo
        {
            CursoRaw    = cursoRaw.Trim(),
            NombreCorto = ExtractShortName(cursoRaw),
            AsuntoEmail = $"Información sobre: {ExtractShortName(cursoRaw)}"
        };
        _data.Courses.Add(course);
        Save();
        return course;
    }

    public void UpdateCourse(CourseInfo updated)
    {
        var existing = _data.Courses.FirstOrDefault(c => c.Id == updated.Id);
        if (existing == null) return;
        CopyTo(updated, existing);
        existing.FechaModificacion = DateTime.Now;
        Save();
    }

    public bool DeleteCourse(string courseId)
    {
        var existing = _data.Courses.FirstOrDefault(c => c.Id == courseId);
        if (existing == null) return false;
        _data.Courses.Remove(existing);
        Save();
        return true;
    }

    // ── Helpers envíos ────────────────────────────────────────────────────────

    public string BuildLeadKey(Lead lead)
    {
        static string N(string? v) => (v ?? string.Empty).Trim().ToLowerInvariant();

        return string.Join("|",
            lead.SourceRow.ToString(),
            N(lead.FechaRegistro),
            N(lead.Curso),
            N(lead.Plataforma),
            N(lead.SituacionLaboral),
            N(lead.NivelEstudios),
            N(lead.Nombre),
            N(lead.Email),
            N(lead.Telefono),
            N(lead.Provincia),
            lead.AceptaPublicidad ? "1" : "0",
            N(lead.Observaciones),
            N(lead.Contacto),
            N(lead.Inscripcion));
    }

    public bool HasBeenSent(Lead lead)
    {
        var key = BuildLeadKey(lead);
        return _data.SentRecords.Any(r =>
            !string.IsNullOrWhiteSpace(r.LeadKey) &&
            string.Equals(r.LeadKey, key, StringComparison.Ordinal));
    }

    public bool HasBeenSentForCourse(Lead lead, string? courseRaw)
    {
        if (string.IsNullOrWhiteSpace(courseRaw))
            return false;

        var key = BuildLeadKey(lead);
        var targetCourse = courseRaw.Trim();

        return _data.SentRecords.Any(r =>
            r.Success &&
            !string.IsNullOrWhiteSpace(r.LeadKey) &&
            string.Equals(r.LeadKey, key, StringComparison.Ordinal) &&
            string.Equals((r.CursoRaw ?? string.Empty).Trim(), targetCourse, StringComparison.OrdinalIgnoreCase));
    }

    public void AddSentRecord(SentRecord record)
    {
        _data.SentRecords.Add(record);
        Save();
    }

    public DateTime? GetLeadEmailSentAt(string key)
        => _data.LeadEmailSentAt.TryGetValue(key, out var sentAt) ? sentAt : null;

    public void SetLeadEmailSentAt(string key, DateTime sentAt)
    {
        _data.LeadEmailSentAt[key] = sentAt;
        Save();
    }

    public DateTime? GetLeadWhatsAppSentAt(string key)
        => _data.LeadWhatsAppSentAt.TryGetValue(key, out var sentAt) ? sentAt : null;

    public void SetLeadWhatsAppSentAt(string key, DateTime sentAt)
    {
        _data.LeadWhatsAppSentAt[key] = sentAt;
        Save();
    }

    public DateTime? GetLeadPhoneCalledAt(string key)
        => _data.LeadPhoneCalledAt.TryGetValue(key, out var calledAt) ? calledAt : null;

    public void SetLeadPhoneCalledAt(string key, DateTime? calledAt)
    {
        if (calledAt.HasValue)
            _data.LeadPhoneCalledAt[key] = calledAt.Value;
        else
            _data.LeadPhoneCalledAt.Remove(key);

        Save();
    }

    public DateTime? GetLastSuccessfulEmailSentAt(string leadKey, string? courseRaw = null)
    {
        var records = _data.SentRecords
            .Where(r => r.Success
                && !string.IsNullOrWhiteSpace(r.LeadKey)
                && string.Equals(r.LeadKey, leadKey, StringComparison.Ordinal));

        if (!string.IsNullOrWhiteSpace(courseRaw))
        {
            var targetCourse = courseRaw.Trim();
            records = records.Where(r => string.Equals((r.CursoRaw ?? string.Empty).Trim(), targetCourse, StringComparison.OrdinalIgnoreCase));
        }

        return records
            .OrderByDescending(r => r.FechaEnvio)
            .Select(r => (DateTime?)r.FechaEnvio)
            .FirstOrDefault();
    }

    // ── Helpers estado de lead ────────────────────────────────────────────────

    public LeadStatus GetLeadStatus(string key)
        => _data.LeadStatusOverrides.TryGetValue(key, out var s) ? s : LeadStatus.Pendiente;

    public void SetLeadStatus(string key, LeadStatus status)
    {
        _data.LeadStatusOverrides[key] = status;
        Save();
    }

    public void RemoveLeadStatus(string key)
    {
        if (_data.LeadStatusOverrides.Remove(key)) Save();
    }

    // ── Helpers etiquetas de lead ─────────────────────────────────────────────

    public LeadLabel GetLeadLabel(string key)
    {
        if (_data.LeadLabels.TryGetValue(key, out var label)) return label;
        // Migración: datos previos guardados como Descartado en LeadStatusOverrides
        if (_data.LeadStatusOverrides.TryGetValue(key, out var status) && status != LeadStatus.Pendiente)
            return LeadLabel.DescartadoNoInteresa;
        return LeadLabel.Ninguna;
    }

    public void SetLeadLabel(string key, LeadLabel label)
    {
        if (label == LeadLabel.Ninguna)
        {
            _data.LeadLabels.Remove(key);
            _data.LeadStatusOverrides.Remove(key);
        }
        else
        {
            _data.LeadLabels[key] = label;
            _data.LeadStatusOverrides.Remove(key);
        }
        Save();
    }

    public void RemoveLeadLabel(string key)
    {
        bool changed = _data.LeadLabels.Remove(key);
        changed |= _data.LeadStatusOverrides.Remove(key);
        if (changed) Save();
    }

    // ── Helpers notas de lead ─────────────────────────────────────────────────

    public string GetLeadNote(string key)
        => _data.LeadNotes.TryGetValue(key, out var note) ? note : string.Empty;

    public void SetLeadNote(string key, string? note)
    {
        var text = (note ?? string.Empty).Trim();
        if (string.IsNullOrEmpty(text))
            _data.LeadNotes.Remove(key);
        else
            _data.LeadNotes[key] = text;

        Save();
    }

    public void RemoveLeadNote(string key)
    {
        if (_data.LeadNotes.Remove(key)) Save();
    }

    // ── Helpers próximo contacto ───────────────────────────────────────────────

    public string GetLeadNextContact(string key)
        => _data.LeadNextContacts.TryGetValue(key, out var next) ? next : string.Empty;

    public string GetLeadHandledBy(string key)
        => _data.LeadHandledBy.TryGetValue(key, out var handledBy) ? handledBy : string.Empty;

    public void SetLeadHandledBy(string key, string? handledBy)
    {
        var value = (handledBy ?? string.Empty).Trim();
        if (string.IsNullOrEmpty(value))
            _data.LeadHandledBy.Remove(key);
        else
            _data.LeadHandledBy[key] = value;

        Save();
    }

    public void SetLeadNextContact(string key, string? next)
    {
        var text = (next ?? string.Empty).Trim();
        if (string.IsNullOrEmpty(text))
            _data.LeadNextContacts.Remove(key);
        else
            _data.LeadNextContacts[key] = text;

        Save();
    }

    public void RemoveLeadNextContact(string key)
    {
        if (_data.LeadNextContacts.Remove(key)) Save();
    }

    // ── Utils ─────────────────────────────────────────────────────────────────

    private static string ExtractShortName(string raw)
    {
        // Toma la parte antes del primer punto o barra
        var name = raw.Split('.')[0].Trim();
        // Elimina prefijos tipo "MD " o "SG "
        if (name.Length > 3 && char.IsLetter(name[0]) && char.IsLetter(name[1]) && name[2] == ' ')
            name = name[3..];
        return name.Length > 80 ? name[..80] : name;
    }

    private static void CopyTo(CourseInfo src, CourseInfo dst)
    {
        dst.NombreCorto            = src.NombreCorto;
        dst.NombreComplementario   = src.NombreComplementario;
        dst.AsuntoEmail            = src.AsuntoEmail;
        dst.TextoMarketing         = src.TextoMarketing;
        dst.TextoMarketingRich     = src.TextoMarketingRich;
        dst.TextoInicioCurso       = src.TextoInicioCurso;
        dst.TextoWhatsApp          = src.TextoWhatsApp;
        dst.TextoSocial            = src.TextoSocial;
        dst.RequisitosAcceso       = src.RequisitosAcceso;
        dst.DocumentacionNecesaria = src.DocumentacionNecesaria;
        dst.FechaInicio            = src.FechaInicio;
        dst.FechaFin               = src.FechaFin;
        dst.HorarioInfo            = src.HorarioInfo;
        dst.UrlFichaInscripcion    = src.UrlFichaInscripcion;
        dst.PdfAdjuntoPath         = src.PdfAdjuntoPath;
        dst.InfoAdicional          = src.InfoAdicional;
    }

    private static AppData NormalizeData(AppData data)
    {
        data.SmtpConfig ??= new SmtpConfig();
        data.SmtpConfig.SendDelayMs = Math.Clamp(data.SmtpConfig.SendDelayMs, 0, 10000);
        data.SmtpConfig.MaxSendsPerSession = Math.Max(0, data.SmtpConfig.MaxSendsPerSession);
        data.SmtpConfig.MaxSendsPerDay = Math.Max(0, data.SmtpConfig.MaxSendsPerDay);
        data.Courses ??= new List<CourseInfo>();
        data.Students ??= new List<Student>();
        data.SentRecords ??= new List<SentRecord>();
        data.LeadStatusOverrides ??= new Dictionary<string, LeadStatus>();
        data.LeadLabels ??= new Dictionary<string, LeadLabel>();
        data.LeadNotes ??= new Dictionary<string, string>();
        data.LeadNextContacts ??= new Dictionary<string, string>();
        data.LeadHandledBy ??= new Dictionary<string, string>();
        data.LeadEmailSentAt ??= new Dictionary<string, DateTime>();
        data.LeadWhatsAppSentAt ??= new Dictionary<string, DateTime>();
        data.LeadPhoneCalledAt ??= new Dictionary<string, DateTime>();
        return data;
    }
}
