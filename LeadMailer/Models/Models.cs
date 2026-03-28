using CommunityToolkit.Mvvm.ComponentModel;

namespace LeadMailer.Models;

// ─── Estado de un lead ───────────────────────────────────────────────────────
public enum LeadStatus { Pendiente, Enviado, Descartado }
public enum LeadLabel  { Ninguna, Interesado, Confirmado, Descartado }

// ─── Lead (una fila del Excel) ───────────────────────────────────────────────
public class Lead
{
    public int    SourceRow        { get; set; }
    public string FechaRegistro    { get; set; } = "";
    public string Curso            { get; set; } = "";
    public string Plataforma       { get; set; } = "";
    public string SituacionLaboral { get; set; } = "";
    public string NivelEstudios    { get; set; } = "";
    public string Nombre           { get; set; } = "";
    public string Email            { get; set; } = "";
    public string Telefono         { get; set; } = "";
    public string Provincia        { get; set; } = "";
    public bool   AceptaPublicidad { get; set; }
    public string Observaciones    { get; set; } = "";
    public string Contacto         { get; set; } = "";
    public string Inscripcion      { get; set; } = "";
}

// ─── LeadRow (Lead + estado de UI) ──────────────────────────────────────────
public partial class LeadRow : ObservableObject
{
    public Lead Lead { get; }
    public CourseInfo Course { get; }

    [ObservableProperty] private bool _isSelected;
    [ObservableProperty] private bool _alreadySent;
    [ObservableProperty] private bool _isDuplicate;
    [ObservableProperty] private bool _isInvalidEmail;
    [ObservableProperty] private LeadLabel _label = LeadLabel.Ninguna;
    [ObservableProperty] private string _note = "";
    [ObservableProperty] private string _nextContact = "";

    public string Nombre => Lead.Nombre;
    public string Email => Lead.Email;
    public string Telefono => Lead.Telefono;
    public string Provincia => Lead.Provincia;
    public string SituacionLaboral => Lead.SituacionLaboral;
    public string Plataforma => Lead.Plataforma;
    public string NombreCurso => Course.NombreCorto;

    public bool IsDiscarded => Label == LeadLabel.Descartado;
    public bool CanSelect => !AlreadySent && Label != LeadLabel.Descartado;

    public LeadStatus Status
    {
        get
        {
            if (IsDiscarded) return LeadStatus.Descartado;
            if (AlreadySent) return LeadStatus.Enviado;
            return LeadStatus.Pendiente;
        }
    }

    public string EstadoTexto => Status switch
    {
        LeadStatus.Enviado    => "Enviado",
        LeadStatus.Descartado => "Descartado",
        _                     => "Pendiente"
    };

    private readonly Action<string>?    _onNoteChanged;
    private readonly Action<string>?    _onNextContactChanged;
    private readonly Action<LeadLabel>? _onLabelChanged;

    public LeadRow(
        Lead lead,
        CourseInfo course,
        bool alreadySent,
        LeadLabel label,
        string note,
        Action<string>? onNoteChanged,
        string nextContact,
        Action<string>? onNextContactChanged,
        Action<LeadLabel>? onLabelChanged = null)
    {
        Lead = lead;
        Course = course;
        _alreadySent = alreadySent;
        _label = label;
        _note = note;
        _onNoteChanged = onNoteChanged;
        _nextContact = nextContact;
        _onNextContactChanged = onNextContactChanged;
        _onLabelChanged = onLabelChanged;
    }

    partial void OnAlreadySentChanged(bool value)
    {
        OnPropertyChanged(nameof(CanSelect));
        OnPropertyChanged(nameof(Status));
        OnPropertyChanged(nameof(EstadoTexto));
    }

    partial void OnLabelChanged(LeadLabel value)
    {
        OnPropertyChanged(nameof(IsDiscarded));
        OnPropertyChanged(nameof(CanSelect));
        OnPropertyChanged(nameof(Status));
        OnPropertyChanged(nameof(EstadoTexto));
        _onLabelChanged?.Invoke(value);
    }

    partial void OnNoteChanged(string value) => _onNoteChanged?.Invoke(value);
    partial void OnNextContactChanged(string value) => _onNextContactChanged?.Invoke(value);
}

// ─── Información de un curso ─────────────────────────────────────────────────
public class CourseInfo
{
    public string   Id                     { get; set; } = Guid.NewGuid().ToString();
    public string   CursoRaw               { get; set; } = "";   // Valor exacto del Excel
    public string   NombreCorto            { get; set; } = "";   // Nombre de display
    public string   NombreComplementario   { get; set; } = "";   // Nombre alternativo para mostrar en email
    public string   AsuntoEmail            { get; set; } = "";
    public string   TextoMarketing         { get; set; } = "";
    public string   TextoMarketingRich     { get; set; } = "";
    public string   TextoWhatsApp          { get; set; } = "";   // Nuevo: Texto para WhatsApp
    public string   RequisitosAcceso       { get; set; } = "";
    public string   DocumentacionNecesaria { get; set; } = "";
    public string   FechaInicio            { get; set; } = "";
    public string   FechaFin               { get; set; } = "";
    public string   HorarioInfo            { get; set; } = "";
    public string   UrlFichaInscripcion    { get; set; } = "";
    public string   PdfAdjuntoPath         { get; set; } = "";
    public string   InfoAdicional          { get; set; } = "";
    public DateTime FechaCreacion          { get; set; } = DateTime.Now;
    public DateTime FechaModificacion      { get; set; } = DateTime.Now;
}

// ─── Config SMTP ─────────────────────────────────────────────────────────────
public class SmtpConfig
{
    public string Host          { get; set; } = "smtp.gmail.com";
    public int    Port          { get; set; } = 587;
    public bool   EnableSsl     { get; set; } = true;
    public string Username      { get; set; } = "";
    public string Password      { get; set; } = "";
    public string FromName      { get; set; } = "";
    public string FromEmail     { get; set; } = "";
    public string ExcelFilePath  { get; set; } = "";
    public string GeminiApiKey     { get; set; } = "";
    public string OpenRouterApiKey { get; set; } = "";
    public string AiProvider       { get; set; } = "OpenRouter";
    public int    SendDelayMs      { get; set; } = 500;
    public string WhatsAppNumber   { get; set; } = "";

    public bool   ConfirmMassSend           { get; set; } = true;
    public bool   EnableScheduledSend       { get; set; }
    public DateTime? ScheduledSendAt        { get; set; }
    public int    MaxSendsPerSession        { get; set; } = 0;
    public int    MaxSendsPerDay            { get; set; } = 0;

    public string LabelTemplateInteresado   { get; set; } = "";
    public string LabelTemplateConfirmado   { get; set; } = "";
    public string LabelTemplateDescartado   { get; set; } = "";
}

// ─── Raíz de persistencia JSON ───────────────────────────────────────────────
public class AppData
{
    public SmtpConfig       SmtpConfig  { get; set; } = new();
    public List<CourseInfo> Courses     { get; set; } = new();
    public List<SentRecord> SentRecords { get; set; } = new();
    /// <summary>LeadKey → estado persistido (solo se almacena si difiere de Pendiente).</summary>
    public Dictionary<string, LeadStatus> LeadStatusOverrides { get; set; } = new();
    /// <summary>LeadKey → etiqueta manual del lead.</summary>
    public Dictionary<string, LeadLabel> LeadLabels { get; set; } = new();
    /// <summary>LeadKey → nota comercial/seguimiento.</summary>
    public Dictionary<string, string> LeadNotes { get; set; } = new();
    /// <summary>LeadKey → próximo contacto (texto libre: fecha/hora/acción).</summary>
    public Dictionary<string, string> LeadNextContacts { get; set; } = new();
}

// ─── Registro de envíos ──────────────────────────────────────────────────────
public class SentRecord
{
    public string   LeadKey    { get; set; } = "";
    public string?  NombreLead { get; set; }
    public string   Email      { get; set; } = "";
    public string   CursoRaw   { get; set; } = "";
    public DateTime FechaEnvio { get; set; } = DateTime.Now;
    public bool     Success    { get; set; }
    public string?  Error      { get; set; }
}
