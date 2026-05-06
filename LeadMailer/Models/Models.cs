using CommunityToolkit.Mvvm.ComponentModel;

namespace LeadMailer.Models;

// ─── Estado de un lead ───────────────────────────────────────────────────────
public enum LeadStatus { Pendiente, EmailEnviado, WhatsAppEnviado, AmbosEnviados }
public enum LeadLabel
{
    Ninguna = 0,
    Solicitado = 1,
    ConfirmaSi = 2,
    DescartadoNoInteresa = 3,
    InscritoPruebaNivel = 4,
    Simultaneidad = 5,
    Reserva = 6,
    DescartadoSuperaHorasCursoCerrado = 7,
    DescartadoColectivo = 8,
    DescartadoPorSector = 9,
    Inscrito = 10,
    Realizado = 11,
    NoIniciaConectaAsiste = 12,
    DescartadoPorTitulacion = 13,
    NoLocalizado = 14,
    BajaLopd = 15,
    DescartadoPruebaCompetencia = 16,
    Erroneo = 17,

    Interesado = Solicitado,
    Confirmado = ConfirmaSi,
    Descartado = DescartadoNoInteresa
}

public static class LeadLabelExtensions
{
    public static bool IsDiscardLabel(this LeadLabel label)
        => label is LeadLabel.DescartadoNoInteresa
            or LeadLabel.DescartadoSuperaHorasCursoCerrado
            or LeadLabel.DescartadoColectivo
            or LeadLabel.DescartadoPorSector
            or LeadLabel.DescartadoPorTitulacion
            or LeadLabel.BajaLopd
            or LeadLabel.DescartadoPruebaCompetencia
            or LeadLabel.Erroneo;
}

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

// ─── Alumno (registro manual) ───────────────────────────────────────────────
public class Student
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
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

// ─── Alumno en UI (manual o lead confirmado) ───────────────────────────────
public class StudentEntry : ObservableObject
{
    private readonly Student? _model;
    private readonly LeadRow? _leadRow;
    private readonly Action? _onChanged;

    private string _fechaRegistro = "";
    private string _curso = "";
    private string _plataforma = "";
    private string _situacionLaboral = "";
    private string _nivelEstudios = "";
    private string _nombre = "";
    private string _email = "";
    private string _telefono = "";
    private string _provincia = "";
    private bool   _aceptaPublicidad;
    private string _observaciones = "";
    private string _contacto = "";
    private string _inscripcion = "";

    public bool IsLinkedToLead => _leadRow != null;
    public bool IsManual => _model != null;
    public string Origin => IsLinkedToLead ? "Lead confirmado" : "Manual";
    public Student? Model => _model;
    public LeadRow? LeadRow => _leadRow;

    public string FechaRegistro
    {
        get => _fechaRegistro;
        set
        {
            if (!SetProperty(ref _fechaRegistro, value)) return;
            if (_model == null) return;
            _model.FechaRegistro = value;
            _onChanged?.Invoke();
        }
    }

    public string Curso
    {
        get => _curso;
        set
        {
            if (!SetProperty(ref _curso, value)) return;
            if (_model != null)
            {
                _model.Curso = value;
                _onChanged?.Invoke();
            }
            OnPropertyChanged(nameof(CursoDisplay));
        }
    }

    public string Plataforma
    {
        get => _plataforma;
        set
        {
            if (!SetProperty(ref _plataforma, value)) return;
            if (_model == null) return;
            _model.Plataforma = value;
            _onChanged?.Invoke();
        }
    }

    public string SituacionLaboral
    {
        get => _situacionLaboral;
        set
        {
            if (!SetProperty(ref _situacionLaboral, value)) return;
            if (_model == null) return;
            _model.SituacionLaboral = value;
            _onChanged?.Invoke();
        }
    }

    public string NivelEstudios
    {
        get => _nivelEstudios;
        set
        {
            if (!SetProperty(ref _nivelEstudios, value)) return;
            if (_model == null) return;
            _model.NivelEstudios = value;
            _onChanged?.Invoke();
        }
    }

    public string Nombre
    {
        get => _nombre;
        set
        {
            if (!SetProperty(ref _nombre, value)) return;
            if (_model == null) return;
            _model.Nombre = value;
            _onChanged?.Invoke();
        }
    }

    public string Email
    {
        get => _email;
        set
        {
            if (!SetProperty(ref _email, value)) return;
            if (_model == null) return;
            _model.Email = value;
            _onChanged?.Invoke();
        }
    }

    public string Telefono
    {
        get => _telefono;
        set
        {
            if (!SetProperty(ref _telefono, value)) return;
            if (_model == null) return;
            _model.Telefono = value;
            _onChanged?.Invoke();
        }
    }

    public string Provincia
    {
        get => _provincia;
        set
        {
            if (!SetProperty(ref _provincia, value)) return;
            if (_model == null) return;
            _model.Provincia = value;
            _onChanged?.Invoke();
        }
    }

    public bool AceptaPublicidad
    {
        get => _aceptaPublicidad;
        set
        {
            if (!SetProperty(ref _aceptaPublicidad, value)) return;
            if (_model == null) return;
            _model.AceptaPublicidad = value;
            _onChanged?.Invoke();
        }
    }

    public string Observaciones
    {
        get => _observaciones;
        set
        {
            if (!SetProperty(ref _observaciones, value)) return;
            if (_model == null) return;
            _model.Observaciones = value;
            _onChanged?.Invoke();
        }
    }

    public string Contacto
    {
        get => _contacto;
        set
        {
            if (!SetProperty(ref _contacto, value)) return;
            if (_model == null) return;
            _model.Contacto = value;
            _onChanged?.Invoke();
        }
    }

    public string Inscripcion
    {
        get => _inscripcion;
        set
        {
            if (!SetProperty(ref _inscripcion, value)) return;
            if (_model == null) return;
            _model.Inscripcion = value;
            _onChanged?.Invoke();
        }
    }

    public string CursoDisplay => IsLinkedToLead ? _leadRow!.NombreCurso : Curso;
    public string CursoRaw => _leadRow?.Lead.Curso ?? Curso;

    public StudentEntry(LeadRow leadRow)
    {
        _leadRow = leadRow;
        _fechaRegistro = leadRow.Lead.FechaRegistro;
        _curso = leadRow.Lead.Curso;
        _plataforma = leadRow.Lead.Plataforma;
        _situacionLaboral = leadRow.Lead.SituacionLaboral;
        _nivelEstudios = leadRow.Lead.NivelEstudios;
        _nombre = leadRow.Lead.Nombre;
        _email = leadRow.Lead.Email;
        _telefono = leadRow.Lead.Telefono;
        _provincia = leadRow.Lead.Provincia;
        _aceptaPublicidad = leadRow.Lead.AceptaPublicidad;
        _observaciones = leadRow.Lead.Observaciones;
        _contacto = leadRow.Lead.Contacto;
        _inscripcion = leadRow.Lead.Inscripcion;
    }

    public StudentEntry(Student model, Action? onChanged)
    {
        _model = model;
        _onChanged = onChanged;
        _fechaRegistro = model.FechaRegistro;
        _curso = model.Curso;
        _plataforma = model.Plataforma;
        _situacionLaboral = model.SituacionLaboral;
        _nivelEstudios = model.NivelEstudios;
        _nombre = model.Nombre;
        _email = model.Email;
        _telefono = model.Telefono;
        _provincia = model.Provincia;
        _aceptaPublicidad = model.AceptaPublicidad;
        _observaciones = model.Observaciones;
        _contacto = model.Contacto;
        _inscripcion = model.Inscripcion;
    }
}

// ─── LeadRow (Lead + estado de UI) ──────────────────────────────────────────
public partial class LeadRow : ObservableObject
{
    public Lead Lead { get; }
    public CourseInfo Course { get; }

    [ObservableProperty] private bool _isSelected;
    [ObservableProperty] private bool _alreadySent;
    [ObservableProperty] private DateTime? _emailSentAt;
    [ObservableProperty] private DateTime? _whatsAppSentAt;
    [ObservableProperty] private bool _isDuplicate;
    [ObservableProperty] private bool _isInvalidEmail;
    [ObservableProperty] private LeadLabel _label = LeadLabel.Ninguna;
    [ObservableProperty] private string _note = "";
    [ObservableProperty] private string _nextContact = "";
    [ObservableProperty] private string _handledBy = "";
    [ObservableProperty] private DateTime? _phoneCalledAt;

    public string Nombre => Lead.Nombre;
    public string Email => Lead.Email;
    public string Telefono => Lead.Telefono;
    public string Provincia => Lead.Provincia;
    public string SituacionLaboral => Lead.SituacionLaboral;
    public string Plataforma => Lead.Plataforma;
    public string NombreCurso => Course.NombreCorto;
    public string PhoneCalledAtText => PhoneCalledAt?.ToString("dd/MM HH:mm") ?? string.Empty;
    public bool PhoneCalled
    {
        get => PhoneCalledAt.HasValue;
        set
        {
            if (value)
            {
                if (!PhoneCalledAt.HasValue)
                    PhoneCalledAt = DateTime.Now;
            }
            else
            {
                if (PhoneCalledAt.HasValue)
                    PhoneCalledAt = null;
            }
        }
    }

    public bool IsDiscarded => Label.IsDiscardLabel();
    public bool CanSelect => !AlreadySent && !IsDiscarded;

    public LeadStatus Status
    {
        get
        {
            if (EmailSentAt.HasValue && WhatsAppSentAt.HasValue) return LeadStatus.AmbosEnviados;
            if (EmailSentAt.HasValue) return LeadStatus.EmailEnviado;
            if (WhatsAppSentAt.HasValue) return LeadStatus.WhatsAppEnviado;
            return LeadStatus.Pendiente;
        }
    }

    public string EstadoTexto => Status switch
    {
        LeadStatus.EmailEnviado    => "Email Enviado",
        LeadStatus.WhatsAppEnviado => "WhatsApp Enviado",
        LeadStatus.AmbosEnviados   => "Ambos Enviados",
        _                          => "Pendiente"
    };

    private readonly Action<string>?    _onNoteChanged;
    private readonly Action<string>?    _onNextContactChanged;
    private readonly Action<LeadLabel>? _onLabelChanged;
    private readonly Action<DateTime?>? _onPhoneCalledAtChanged;

    public LeadRow(
        Lead lead,
        CourseInfo course,
        bool alreadySent,
        DateTime? emailSentAt,
        DateTime? whatsAppSentAt,
        DateTime? phoneCalledAt,
        string handledBy,
        LeadLabel label,
        string note,
        Action<string>? onNoteChanged,
        string nextContact,
        Action<string>? onNextContactChanged,
        Action<LeadLabel>? onLabelChanged = null,
        Action<DateTime?>? onPhoneCalledAtChanged = null)
    {
        Lead = lead;
        Course = course;
        _emailSentAt = emailSentAt;
        _whatsAppSentAt = whatsAppSentAt;
        _phoneCalledAt = phoneCalledAt;
        _handledBy = handledBy;
        _alreadySent = alreadySent || emailSentAt.HasValue;
        _label = label;
        _note = note;
        _onNoteChanged = onNoteChanged;
        _nextContact = nextContact;
        _onNextContactChanged = onNextContactChanged;
        _onLabelChanged = onLabelChanged;
        _onPhoneCalledAtChanged = onPhoneCalledAtChanged;
    }

    partial void OnAlreadySentChanged(bool value)
    {
        OnPropertyChanged(nameof(CanSelect));
    }

    partial void OnEmailSentAtChanged(DateTime? value)
    {
        AlreadySent = value.HasValue;
        OnPropertyChanged(nameof(Status));
        OnPropertyChanged(nameof(EstadoTexto));
    }

    partial void OnWhatsAppSentAtChanged(DateTime? value)
    {
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
    partial void OnPhoneCalledAtChanged(DateTime? value)
    {
        OnPropertyChanged(nameof(PhoneCalled));
        OnPropertyChanged(nameof(PhoneCalledAtText));
        _onPhoneCalledAtChanged?.Invoke(value);
    }
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
    public string   TextoInicioCurso       { get; set; } = "";
    public string   TextoWhatsApp          { get; set; } = "";   // Nuevo: Texto para WhatsApp
    public string   TextoSocial            { get; set; } = "";   // Texto para redes sociales
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
    public string CaptadoraNombre { get; set; } = "";
    public string CaptadoraApellidos { get; set; } = "";
    public string CloseReportRecipientEmail { get; set; } = "gerardo.sanz@grupoaspasia.com";
    public string ReportFromDateText { get; set; } = "";
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
    public string EmailFooter               { get; set; } = "";
}

// ─── Raíz de persistencia JSON ───────────────────────────────────────────────
public class AppData
{
    public SmtpConfig       SmtpConfig  { get; set; } = new();
    public List<CourseInfo> Courses     { get; set; } = new();
    public List<Student>    Students    { get; set; } = new();
    public List<SentRecord> SentRecords { get; set; } = new();
    /// <summary>LeadKey → estado persistido (solo se almacena si difiere de Pendiente).</summary>
    public Dictionary<string, LeadStatus> LeadStatusOverrides { get; set; } = new();
    /// <summary>LeadKey → etiqueta manual del lead.</summary>
    public Dictionary<string, LeadLabel> LeadLabels { get; set; } = new();
    /// <summary>LeadKey → nota comercial/seguimiento.</summary>
    public Dictionary<string, string> LeadNotes { get; set; } = new();
    /// <summary>LeadKey → próximo contacto (texto libre: fecha/hora/acción).</summary>
    public Dictionary<string, string> LeadNextContacts { get; set; } = new();
    /// <summary>LeadKey → captadora asignada al tratar el lead.</summary>
    public Dictionary<string, string> LeadHandledBy { get; set; } = new();
    /// <summary>LeadKey → fecha/hora del último envío exitoso por email.</summary>
    public Dictionary<string, DateTime> LeadEmailSentAt { get; set; } = new();
    /// <summary>LeadKey → fecha/hora de apertura de WhatsApp para el lead.</summary>
    public Dictionary<string, DateTime> LeadWhatsAppSentAt { get; set; } = new();
    /// <summary>LeadKey → fecha/hora de llamada telefónica.</summary>
    public Dictionary<string, DateTime> LeadPhoneCalledAt { get; set; } = new();
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
