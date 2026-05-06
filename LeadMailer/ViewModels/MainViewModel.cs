using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using LeadMailer.Helpers;
using LeadMailer.Models;
using LeadMailer.Services;
using Microsoft.Win32;
using OfficeOpenXml;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Windows;
using System.Windows.Data;

namespace LeadMailer.ViewModels
{
    public record LabelOption(LeadLabel Value, string Display);
    public record CourseHistoryMetric(string Curso, int Total, int Correctos, int Fallidos, double SuccessRate);

    public partial class MainViewModel : ObservableObject
    {
        private readonly DataService   _data;
        private readonly ExcelService  _email;
        private readonly EmailService  _emailService;
        private readonly AiTextService _aiText;

        // Centinela "Todos los cursos" para el ComboBox de filtro
        private readonly CourseInfo _todosCursos = new() { Id = "", NombreCorto = "— Todos los cursos —" };

        [ObservableProperty] private string _currentView = "Leads";
        [ObservableProperty] private ObservableCollection<LeadRow> _leads = new();
        [ObservableProperty] private string _excelFilePath   = "";
        [ObservableProperty] private string _statusMessage   = "";
        [ObservableProperty] private bool   _isBusy;
        [ObservableProperty] private bool   _hasCourseWarning;
        [ObservableProperty] private string _courseWarningText = "";
        [ObservableProperty] private int    _pendingCount;
        [ObservableProperty] private int    _sentCount;
        [ObservableProperty] private int    _followUpTodayCount;
        [ObservableProperty] private int    _filteredCount;
        [ObservableProperty] private int _leadIssuesTotalCount;
        [ObservableProperty] private int _leadIssuesDuplicateCount;
        [ObservableProperty] private int _leadIssuesInvalidEmailCount;
        [ObservableProperty] private int _leadIssuesUnconfiguredCourseCount;
        [ObservableProperty] private ObservableCollection<CourseInfo> _courses = new();
        [ObservableProperty] private ObservableCollection<CourseInfo> _courseFilterItems = new();
        [ObservableProperty] private CourseInfo  _selectedCourseFilter = new();
        [ObservableProperty] private string      _selectedStatusFilter = "Todos";
        [ObservableProperty] private string      _searchText = "";
        [ObservableProperty] private CourseInfo? _editingCourse;
        [ObservableProperty] private bool   _isGenerating;
        [ObservableProperty] private string _geminiError = "";
        [ObservableProperty] private bool   _hasGeminiError;
        [ObservableProperty] private SmtpConfig _smtpConfig = new();
        [ObservableProperty] private string     _smtpTestResult = "";

        [ObservableProperty] private ObservableCollection<SentRecord> _sentHistory = new();
        [ObservableProperty] private string _historySearchText = "";
        [ObservableProperty] private string _historyStatusFilter = "Todos";
        [ObservableProperty] private DateTime? _historyFromDate;
        [ObservableProperty] private DateTime? _historyToDate;
        [ObservableProperty] private int _historySuccessCount;
        [ObservableProperty] private int _historyFailedCount;
        [ObservableProperty] private int _historyTodayCount;
        [ObservableProperty] private int _historyLast7DaysCount;
        [ObservableProperty] private int _historyLast30DaysCount;
        [ObservableProperty] private ObservableCollection<CourseHistoryMetric> _historyCourseMetrics = new();
        [ObservableProperty] private string _whatsappMessage = "";
        [ObservableProperty] private string _socialMessage = "";
        [ObservableProperty] private string _courseStartMessage = "";

        [ObservableProperty] private ObservableCollection<StudentEntry> _students = new();
        [ObservableProperty] private string _studentSearchText = "";
        [ObservableProperty] private ObservableCollection<string> _studentCourseFilterOptions = new();
        [ObservableProperty] private string _selectedStudentCourseFilter = "Todos";
        [ObservableProperty] private string _newStudentNombre = "";
        [ObservableProperty] private string _newStudentEmail = "";
        [ObservableProperty] private string _newStudentTelefono = "";
        [ObservableProperty] private string _newStudentCurso = "";
        [ObservableProperty] private string _newStudentPlataforma = "";
        [ObservableProperty] private string _newStudentSituacionLaboral = "";
        [ObservableProperty] private string _newStudentNivelEstudios = "";
        [ObservableProperty] private string _newStudentProvincia = "";
        [ObservableProperty] private string _newStudentObservaciones = "";
        [ObservableProperty] private string _newStudentContacto = "";
        [ObservableProperty] private string _newStudentInscripcion = "";
        [ObservableProperty] private string _newStudentFechaRegistro = "";
        [ObservableProperty] private bool _newStudentAceptaPublicidad = true;
        [ObservableProperty] private bool _isAddStudentDialogOpen;

        [ObservableProperty] private ObservableCollection<LeadRow> _legacyLeads = new();
        [ObservableProperty] private string _legacyExcelFilePath = "";
        [ObservableProperty] private string _legacySearchText = "";
        [ObservableProperty] private string _legacyOriginCourseFilter = "Todos";
        [ObservableProperty] private string _legacyProvinceFilter = "Todas";
        [ObservableProperty] private string _legacyEmploymentFilter = "Todas";
        [ObservableProperty] private bool _legacyOnlyWithPublicidad = true;
        [ObservableProperty] private bool _legacyOnlyDuplicates;
        [ObservableProperty] private bool _legacyOnlyInvalidEmail;
        [ObservableProperty] private bool _legacyOnlyIssues;
        [ObservableProperty] private int _legacyFilteredCount;
        [ObservableProperty] private int _legacySelectedCount;
        [ObservableProperty] private int _legacyIssuesTotalCount;
        [ObservableProperty] private int _legacyIssuesDuplicateCount;
        [ObservableProperty] private int _legacyIssuesInvalidEmailCount;
        [ObservableProperty] private int _legacyIssuesCampaignCourseCount;
        [ObservableProperty] private bool _showLeadsFilters;
        [ObservableProperty] private bool _showLegacyFilters;
        [ObservableProperty] private ObservableCollection<string> _legacyOriginCourseFilterOptions = new();
        [ObservableProperty] private ObservableCollection<string> _legacyProvinceFilterOptions = new();
        [ObservableProperty] private ObservableCollection<string> _legacyEmploymentFilterOptions = new();
        [ObservableProperty] private CourseInfo? _legacySelectedCampaignCourse;

        public bool SmtpTestResultIsError => SmtpTestResult.StartsWith("✗");

        public string AppVersion
        {
            get
            {
                var v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                return v is null ? "v?" : $"v{v.Major}.{v.Minor}.{v.Build}";
            }
        }

        [RelayCommand]
        private void OpenAddStudentDialog()
        {
            IsAddStudentDialogOpen = true;
            if (string.IsNullOrWhiteSpace(NewStudentFechaRegistro))
                NewStudentFechaRegistro = DateTime.Now.ToString("dd/MM/yyyy");

            var selectedCourse = GetSelectedCourseForNewStudent();
            if (!string.IsNullOrWhiteSpace(selectedCourse))
                NewStudentCurso = selectedCourse;
        }

        [RelayCommand]
        private void CancelAddStudentDialog()
        {
            IsAddStudentDialogOpen = false;
            ClearNewStudentFields();
        }

        [RelayCommand]
        private void AddStudent()
        {
            var selectedCourse = GetSelectedCourseForNewStudent();
            var courseToSave = string.IsNullOrWhiteSpace(selectedCourse)
                ? NewStudentCurso
                : selectedCourse;

            if (string.IsNullOrWhiteSpace(NewStudentNombre) || string.IsNullOrWhiteSpace(courseToSave))
            {
                StatusMessage = "Completa al menos nombre y curso del alumno.";
                return;
            }

            var student = new Student
            {
                Nombre = NewStudentNombre.Trim(),
                Email = NewStudentEmail.Trim(),
                Telefono = NewStudentTelefono.Trim(),
                Curso = courseToSave.Trim(),
                Plataforma = NewStudentPlataforma.Trim(),
                SituacionLaboral = NewStudentSituacionLaboral.Trim(),
                NivelEstudios = NewStudentNivelEstudios.Trim(),
                Provincia = NewStudentProvincia.Trim(),
                Observaciones = NewStudentObservaciones.Trim(),
                Contacto = NewStudentContacto.Trim(),
                Inscripcion = NewStudentInscripcion.Trim(),
                FechaRegistro = string.IsNullOrWhiteSpace(NewStudentFechaRegistro)
                    ? DateTime.Now.ToString("dd/MM/yyyy")
                    : NewStudentFechaRegistro.Trim(),
                AceptaPublicidad = NewStudentAceptaPublicidad
            };

            _data.Students.Add(student);
            _data.Save();
            RefreshStudents();
            ClearNewStudentFields();
            IsAddStudentDialogOpen = false;
            StatusMessage = $"✓  Alumno '{student.Nombre}' añadido.";
        }

        private string GetSelectedCourseForNewStudent()
        {
            if (string.IsNullOrWhiteSpace(SelectedStudentCourseFilter)
                || SelectedStudentCourseFilter == "Todos")
                return string.Empty;

            return SelectedStudentCourseFilter.Trim();
        }

        [RelayCommand]
        private void DeleteStudent(StudentEntry? row)
        {
            if (row?.Model == null)
            {
                StatusMessage = "No se puede eliminar un alumno vinculado a un lead confirmado.";
                return;
            }

            var confirm = MessageBox.Show(
                $"¿Seguro que quieres eliminar al alumno '{row.Nombre}'?",
                "Confirmar eliminación", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;

            _data.Students.Remove(row.Model);
            _data.Save();
            RefreshStudents();
            StatusMessage = $"✓  Alumno '{row.Nombre}' eliminado.";
        }

        [RelayCommand]
        private async Task SendStudentsEmails()
        {
            var students = (FilteredStudents?.Cast<StudentEntry>().ToList() ?? Students.ToList());
            if (students.Count == 0)
            {
                StatusMessage = "No hay alumnos filtrados para avisar por email.";
                return;
            }

            var candidates = students
                .Where(s => !string.IsNullOrWhiteSpace(s.Email))
                .ToList();

            if (candidates.Count == 0)
            {
                StatusMessage = "No hay alumnos con email para avisar.";
                return;
            }

            if (!SmtpOk()) return;

            if (_data.SmtpConfig.ConfirmMassSend && candidates.Count > 1)
            {
                var confirm = MessageBox.Show(
                    $"Vas a enviar {candidates.Count} emails de aviso. ¿Deseas continuar?",
                    "Confirmar envío masivo",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                if (confirm != MessageBoxResult.Yes) return;
            }

            var scheduledAt = _data.SmtpConfig.EnableScheduledSend ? _data.SmtpConfig.ScheduledSendAt : null;
            if (scheduledAt is DateTime when && when > DateTime.Now)
            {
                var waitMs = (int)Math.Min((when - DateTime.Now).TotalMilliseconds, int.MaxValue);
                StatusMessage = $"Aviso programado para {when:dd/MM/yyyy HH:mm}. Esperando…";
                await Task.Delay(waitMs);
            }

            var originalCount = candidates.Count;
            candidates = ApplyStudentSendLimits(candidates, "alumnos");
            if (candidates.Count == 0)
            {
                StatusMessage = "No se puede iniciar el aviso por límites de envío configurados.";
                return;
            }

            IsBusy = true;
            int ok = 0, fail = 0;
            var failDetails = new List<string>();

            for (int i = 0; i < candidates.Count; i++)
            {
                var student = candidates[i];
                StatusMessage = $"Aviso alumnos: enviando {i + 1} de {candidates.Count}: {student.Email}…";

                var course = GetCourseForStudent(student);
                if (course == null)
                {
                    fail++;
                    failDetails.Add($"{student.Email}: curso no configurado.");
                    continue;
                }

                if (!IsCourseConfigured(course))
                {
                    fail++;
                    failDetails.Add($"{student.Email}: curso sin contenido configurado.");
                    continue;
                }

                var (success, error) = await SendStudentWithCourseAsync(student, course);
                if (success) ok++;
                else
                {
                    fail++;
                    var err = string.IsNullOrWhiteSpace(error) ? "Error desconocido al enviar." : error;
                    failDetails.Add($"{student.Email}: {err}");
                }

                var delayMs = Math.Clamp(_data.SmtpConfig.SendDelayMs, 0, 10000);
                if (i < candidates.Count - 1 && delayMs > 0)
                    await Task.Delay(delayMs);
            }

            IsBusy = false;

            if (fail > 0)
            {
                var preview = string.Join(" | ", failDetails.Take(3));
                var more = failDetails.Count > 3 ? $" (+{failDetails.Count - 3} más)" : "";
                StatusMessage = $"Aviso completado — ✓ {ok} enviados | ✗ {fail} fallidos. Detalle: {preview}{more}";
            }
            else
            {
                StatusMessage = $"Aviso completado — ✓ {ok} enviados  |  ✗ {fail} fallidos";
            }

            if (candidates.Count < originalCount)
                StatusMessage += $" | Límite aplicado: {candidates.Count}/{originalCount}";
        }

        [RelayCommand]
        private void SendStudentsWhatsApp()
        {
            var students = (FilteredStudents?.Cast<StudentEntry>().ToList() ?? Students.ToList());
            if (students.Count == 0)
            {
                StatusMessage = "No hay alumnos filtrados para WhatsApp.";
                return;
            }

            int opened = 0, invalid = 0, missingCourse = 0;
            foreach (var student in students)
            {
                var course = GetCourseForStudent(student);
                if (course == null)
                {
                    missingCourse++;
                    continue;
                }

                if (TryOpenWhatsApp(student, course))
                {
                    opened++;
                    if (student.LeadRow != null)
                        MarkLeadWhatsAppOpened(student.LeadRow, DateTime.Now);
                }
                else invalid++;
            }

            var parts = new List<string> { $"WhatsApp alumnos: {opened} chats abiertos" };
            if (invalid > 0) parts.Add($"{invalid} sin teléfono válido");
            if (missingCourse > 0) parts.Add($"{missingCourse} sin curso configurado");
            StatusMessage = string.Join(", ", parts) + ".";
        }

        public IReadOnlyList<string> StatusFilterOptions { get; } =
            new[] { "Todos", "Pendiente", "Email Enviado", "WhatsApp Enviado", "Ambos Enviados", "Incidencias", "Seguimiento hoy" };

        public IReadOnlyList<LabelOption> LeadLabelOptions { get; } = new LabelOption[]
        {
            new(LeadLabel.Ninguna,    "—"),
            new(LeadLabel.InscritoPruebaNivel, "00 - Inscrito prueba nivel"),
            new(LeadLabel.Solicitado, "01 - Solicitado"),
            new(LeadLabel.Simultaneidad, "02 - Simultaneidad"),
            new(LeadLabel.Reserva, "03 - Reserva"),
            new(LeadLabel.ConfirmaSi, "04 - Confirma SI"),
            new(LeadLabel.DescartadoSuperaHorasCursoCerrado, "05 - Descartado supera horas / Curso cerrado"),
            new(LeadLabel.DescartadoNoInteresa, "06 - Descartado no interesa"),
            new(LeadLabel.DescartadoColectivo, "07 - Descartado colectivo"),
            new(LeadLabel.DescartadoPorSector, "08 - Descartado por sector"),
            new(LeadLabel.Inscrito, "09 - Inscrito"),
            new(LeadLabel.Realizado, "10 - Realizado"),
            new(LeadLabel.NoIniciaConectaAsiste, "11 - No inicia (conecta/asiste)"),
            new(LeadLabel.DescartadoPorTitulacion, "12 - Descartado por titulación"),
            new(LeadLabel.NoLocalizado, "13 - No localizado"),
            new(LeadLabel.BajaLopd, "14 - Baja LOPD"),
            new(LeadLabel.DescartadoPruebaCompetencia, "15 - Descartado Prueba de Competencia"),
            new(LeadLabel.Erroneo, "16 - Erróneo"),
        };

        public IReadOnlyList<string> HistoryStatusFilterOptions { get; } =
            new[] { "Todos", "Correctos", "Fallidos" };

        private ICollectionView? _filteredLeads;
        public ICollectionView FilteredLeads => _filteredLeads!;

        private ICollectionView? _filteredHistory;
        public ICollectionView FilteredHistory => _filteredHistory!;

        private ICollectionView? _filteredLegacyLeads;
        public ICollectionView FilteredLegacyLeads => _filteredLegacyLeads!;

        private ICollectionView? _filteredStudents;
        public ICollectionView FilteredStudents => _filteredStudents!;

        private readonly List<string> _lastFailedLeadKeys = new();
        private Action? _undoAction;
        private string _undoDescription = "";

        public MainViewModel()
        {
            _data   = new DataService();
            _email  = new ExcelService();
            _emailService  = new EmailService();
            _aiText = new AiTextService();
            SmtpConfig    = _data.SmtpConfig;
            ExcelFilePath = _data.SmtpConfig.ExcelFilePath;
            RefreshCourses();
            SelectedCourseFilter = _todosCursos;
            BuildFilteredView();
            BuildHistoryView();
            BuildLegacyFilteredView();
            BuildStudentsView();
            RefreshHistory();
            RefreshStudents();
        }

        // ── Reacción a cambios de filtro ──────────────────────────────────────
        partial void OnSearchTextChanged(string value)              => RefreshFilter();
        partial void OnSelectedCourseFilterChanged(CourseInfo value) => RefreshFilter();
        partial void OnSelectedStatusFilterChanged(string value)    => RefreshFilter();
        partial void OnHistorySearchTextChanged(string value)       => RefreshHistoryFilter();
        partial void OnHistoryStatusFilterChanged(string value)     => RefreshHistoryFilter();
        partial void OnHistoryFromDateChanged(DateTime? value)      => RefreshHistoryFilter();
        partial void OnHistoryToDateChanged(DateTime? value)        => RefreshHistoryFilter();
        partial void OnLegacySearchTextChanged(string value)        => RefreshLegacyFilter();
        partial void OnLegacyOriginCourseFilterChanged(string value) => RefreshLegacyFilter();
        partial void OnLegacyProvinceFilterChanged(string value)    => RefreshLegacyFilter();
        partial void OnLegacyEmploymentFilterChanged(string value)  => RefreshLegacyFilter();
        partial void OnLegacyOnlyWithPublicidadChanged(bool value)  => RefreshLegacyFilter();
        partial void OnLegacyOnlyDuplicatesChanged(bool value)      => RefreshLegacyFilter();
        partial void OnLegacyOnlyInvalidEmailChanged(bool value)    => RefreshLegacyFilter();
        partial void OnLegacyOnlyIssuesChanged(bool value)          => RefreshLegacyFilter();
        partial void OnStudentSearchTextChanged(string value)       => RefreshStudentsFilter();
        partial void OnSelectedStudentCourseFilterChanged(string value) => RefreshStudentsFilter();
        partial void OnLegacySelectedCampaignCourseChanged(CourseInfo? value)
        {
            RefreshLegacySentStateForSelectedCampaignCourse();
            RefreshLegacyFilter();
        }

        [RelayCommand] private void Navigate(string view) => CurrentView = view;

        [RelayCommand]
        private void ToggleLeadsFilters()
            => ShowLeadsFilters = !ShowLeadsFilters;

        [RelayCommand]
        private void ToggleLegacyFilters()
            => ShowLegacyFilters = !ShowLegacyFilters;

        [RelayCommand]
        private void BrowseExcel()
        {
            var dlg = new OpenFileDialog { Filter = "Archivos Excel|*.xlsx;*.xls|Todos|*.*", Title = "Seleccionar archivo Excel de leads" };
            if (dlg.ShowDialog() != true) return;
            ExcelFilePath = dlg.FileName;
            _data.SmtpConfig.ExcelFilePath = dlg.FileName;
            _data.Save();
        }

        [RelayCommand]
        private void LoadLeads()
        {
            if (string.IsNullOrWhiteSpace(ExcelFilePath)) { StatusMessage = "⚠  Selecciona primero el archivo Excel."; return; }
            try
            {
                var rawLeads   = _email.ReadLeads(ExcelFilePath);
                var newCourses = new List<string>();
                Leads.Clear();
                foreach (var lead in rawLeads)
                {
                    var course       = _data.GetOrCreateCourse(lead.Curso);
                    var  key         = _data.BuildLeadKey(lead);
                    var emailSentAt  = _data.GetLeadEmailSentAt(key) ?? _data.GetLastSuccessfulEmailSentAt(key, lead.Curso);
                    var whatsAppSentAt = _data.GetLeadWhatsAppSentAt(key);
                    var phoneCalledAt = _data.GetLeadPhoneCalledAt(key);
                    var handledBy   = _data.GetLeadHandledBy(key);
                    var  label       = _data.GetLeadLabel(key);
                    var note         = _data.GetLeadNote(key);
                    var nextContact  = _data.GetLeadNextContact(key);

                    if (!IsCourseConfigured(course) && !newCourses.Contains(lead.Curso))
                        newCourses.Add(lead.Curso);

                    LeadRow row = null!;
                    row = new LeadRow(
                        lead: lead,
                        course: course,
                        alreadySent: emailSentAt.HasValue,
                        emailSentAt: emailSentAt,
                        whatsAppSentAt: whatsAppSentAt,
                        phoneCalledAt: phoneCalledAt,
                        handledBy: handledBy,
                        label: label,
                        note: note,
                        onNoteChanged: n => _data.SetLeadNote(key, n),
                        nextContact: nextContact,
                        onNextContactChanged: n =>
                        {
                            _data.SetLeadNextContact(key, n);
                            UpdateCounts();
                            RefreshFilter();
                        },
                        onLabelChanged: l =>
                        {
                            _data.SetLeadLabel(key, l);
                            MarkLeadHandled(row);
                            UpdateCounts();
                            RefreshFilter();
                            RefreshStudents();
                        },
                        onPhoneCalledAtChanged: calledAt => _data.SetLeadPhoneCalledAt(key, calledAt));
                    Leads.Add(row);
                }
                RefreshCourses();
                MarkDuplicatesForRows(Leads);
                UpdateCounts();
                RefreshFilter();
                RefreshLeadIssuesSummary();
                RefreshStudents();

                UpdateCourseWarning();
                StatusMessage = rawLeads.Count == 0
                    ? "El archivo no contiene filas de datos."
                    : $"✓  {rawLeads.Count} leads cargados — {PendingCount} pendientes de envío.";
            }
            catch (Exception ex) { StatusMessage = $"Error al leer el Excel: {ex.Message}"; }
        }

        [RelayCommand]
        private void BrowseLegacyExcel()
        {
            var dlg = new OpenFileDialog { Filter = "Archivos Excel|*.xlsx;*.xls|Todos|*.*", Title = "Seleccionar Excel de leads antiguos" };
            if (dlg.ShowDialog() != true) return;
            LegacyExcelFilePath = dlg.FileName;
        }

        [RelayCommand]
        private void LoadLegacyLeads()
        {
            if (string.IsNullOrWhiteSpace(LegacyExcelFilePath))
            {
                StatusMessage = "⚠  Selecciona primero el Excel de leads antiguos.";
                return;
            }

            try
            {
                foreach (var row in LegacyLeads)
                    row.PropertyChanged -= OnLegacyLeadPropertyChanged;

                var rawLeads = _email.ReadLeads(LegacyExcelFilePath);
                LegacyLeads.Clear();

                foreach (var lead in rawLeads)
                {
                    var course = new CourseInfo
                    {
                        Id = $"legacy:{(lead.Curso ?? string.Empty).Trim().ToLowerInvariant()}",
                        CursoRaw = lead.Curso ?? string.Empty,
                        NombreCorto = string.IsNullOrWhiteSpace(lead.Curso) ? "(Sin curso histórico)" : lead.Curso.Trim()
                    };
                    var key = _data.BuildLeadKey(lead);
                    var emailSentAt = _data.GetLastSuccessfulEmailSentAt(key, LegacySelectedCampaignCourse?.CursoRaw);
                    var whatsAppSentAt = _data.GetLeadWhatsAppSentAt(key);
                    var phoneCalledAt = _data.GetLeadPhoneCalledAt(key);
                    var handledBy = _data.GetLeadHandledBy(key);

                    LeadRow row = null!;
                    row = new LeadRow(
                        lead: lead,
                        course: course,
                        alreadySent: emailSentAt.HasValue,
                        emailSentAt: emailSentAt,
                        whatsAppSentAt: whatsAppSentAt,
                        phoneCalledAt: phoneCalledAt,
                        handledBy: handledBy,
                        label: _data.GetLeadLabel(key),
                        note: _data.GetLeadNote(key),
                        onNoteChanged: n => _data.SetLeadNote(key, n),
                        nextContact: _data.GetLeadNextContact(key),
                        onNextContactChanged: n => _data.SetLeadNextContact(key, n),
                        onLabelChanged: l =>
                        {
                            _data.SetLeadLabel(key, l);
                            MarkLeadHandled(row);
                        },
                        onPhoneCalledAtChanged: calledAt => _data.SetLeadPhoneCalledAt(key, calledAt));

                    row.PropertyChanged += OnLegacyLeadPropertyChanged;
                    LegacyLeads.Add(row);
                }

                RefreshCourses();
                MarkDuplicatesForRows(LegacyLeads);
                RefreshLegacySentStateForSelectedCampaignCourse();
                RefreshLegacyFilterOptions();
                RefreshLegacyFilter();
                RefreshLegacySelectionCount();

                StatusMessage = rawLeads.Count == 0
                    ? "El Excel de leads antiguos no contiene filas válidas."
                    : $"✓  {rawLeads.Count} leads antiguos cargados.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error al cargar leads antiguos: {ex.Message}";
            }
        }

        private void RegisterStudentRow(StudentEntry row)
        {
            row.PropertyChanged += OnStudentRowPropertyChanged;
        }

        private void OnStudentRowPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(StudentEntry.Curso)
                || e.PropertyName == nameof(StudentEntry.CursoDisplay))
                RefreshStudentsFilter();
        }

        [RelayCommand]
        private void SelectAllLegacy()
        {
            foreach (var row in FilteredLegacyLeads.Cast<LeadRow>().Where(x => x.CanSelect))
                row.IsSelected = true;
            RefreshLegacySelectionCount();
        }

        [RelayCommand]
        private void DeselectAllLegacy()
        {
            foreach (var row in LegacyLeads)
                row.IsSelected = false;
            RefreshLegacySelectionCount();
        }

        [RelayCommand]
        private async Task SendLegacySelectedEmails()
        {
            if (LegacySelectedCampaignCourse == null)
            {
                StatusMessage = "⚠  Selecciona un curso objetivo para el envío de la campaña.";
                return;
            }

            var selected = LegacyLeads.Where(l => l.IsSelected && !l.IsDiscarded).ToList();
            if (selected.Count == 0)
            {
                StatusMessage = "No hay leads antiguos seleccionados para enviar.";
                return;
            }

            if (!SmtpOk()) return;

            if (_data.SmtpConfig.ConfirmMassSend && selected.Count > 1)
            {
                var confirm = MessageBox.Show(
                    $"Vas a enviar {selected.Count} emails de campaña. ¿Deseas continuar?",
                    "Confirmar envío masivo",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                if (confirm != MessageBoxResult.Yes) return;
            }

            var scheduledAt = _data.SmtpConfig.EnableScheduledSend ? _data.SmtpConfig.ScheduledSendAt : null;
            if (scheduledAt is DateTime when && when > DateTime.Now)
            {
                var waitMs = (int)Math.Min((when - DateTime.Now).TotalMilliseconds, int.MaxValue);
                StatusMessage = $"Campaña programada para {when:dd/MM/yyyy HH:mm}. Esperando…";
                await Task.Delay(waitMs);
            }

            var originalCount = selected.Count;
            selected = ApplySendLimits(selected, "campaña");
            if (selected.Count == 0)
            {
                StatusMessage = "No se puede iniciar la campaña por límites de envío configurados.";
                return;
            }

            IsBusy = true;
            int ok = 0, fail = 0;

            for (int i = 0; i < selected.Count; i++)
            {
                var row = selected[i];
                StatusMessage = $"Campaña: enviando {i + 1} de {selected.Count}: {row.Email}…";

                var (success, _) = await SendLeadWithCourseAsync(
                    row,
                    LegacySelectedCampaignCourse,
                    forceSend: true,
                    recordCourseRaw: LegacySelectedCampaignCourse.CursoRaw);

                if (success) ok++;
                else fail++;

                var delayMs = Math.Clamp(_data.SmtpConfig.SendDelayMs, 0, 10000);
                if (i < selected.Count - 1 && delayMs > 0)
                    await Task.Delay(delayMs);
            }

            IsBusy = false;
            RefreshLegacyFilter();
            UpdateCounts();
            RefreshFilter();

            StatusMessage = $"Campaña completada — ✓ {ok} enviados | ✗ {fail} fallidos";
            if (selected.Count < originalCount)
                StatusMessage += $" | Límite aplicado: {selected.Count}/{originalCount}";
        }

        [RelayCommand]
        private void SendLegacyWhatsAppToSelected()
        {
            if (LegacySelectedCampaignCourse == null)
            {
                StatusMessage = "⚠  Selecciona un curso objetivo para WhatsApp.";
                return;
            }

            var selected = LegacyLeads.Where(l => l.IsSelected).ToList();
            if (selected.Count == 0)
            {
                StatusMessage = "Selecciona al menos un lead antiguo para WhatsApp.";
                return;
            }

            int opened = 0, invalid = 0;
            foreach (var row in selected)
            {
                if (TryOpenWhatsApp(row, LegacySelectedCampaignCourse))
                {
                    opened++;
                    MarkLeadWhatsAppOpened(row, DateTime.Now);
                }
                else invalid++;
            }

            StatusMessage = invalid > 0
                ? $"Campaña WhatsApp: {opened} chats abiertos, {invalid} sin teléfono válido."
                : $"Campaña WhatsApp: {opened} chats abiertos.";
        }

        [RelayCommand]
        private void CopySelectedLegacyEmails()
        {
            var emails = LegacyLeads
                .Where(l => l.IsSelected)
                .Select(l => (l.Email ?? string.Empty).Trim())
                .Where(e => !string.IsNullOrWhiteSpace(e))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (emails.Count == 0)
            {
                StatusMessage = "No hay emails seleccionados en leads antiguos para copiar.";
                return;
            }

            try
            {
                Clipboard.SetText(string.Join(";", emails));
                StatusMessage = $"✓  Copiados {emails.Count} emails de leads antiguos al portapapeles.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"✗  No se pudo copiar al portapapeles: {ex.Message}";
            }
        }

        [RelayCommand]
        private async Task SendSelected()
        {
            var visible = (FilteredLeads?.Cast<LeadRow>().ToList() ?? Leads.ToList());
            var pending = visible.Where(l => l.IsSelected && !l.AlreadySent && !l.IsDiscarded).ToList();
            if (pending.Count == 0) { StatusMessage = "No hay leads seleccionados sin enviar."; return; }
            if (!SmtpOk()) return;

            if (_data.SmtpConfig.ConfirmMassSend && pending.Count > 1)
            {
                var confirm = MessageBox.Show(
                    $"Vas a enviar {pending.Count} emails. ¿Deseas continuar?",
                    "Confirmar envío masivo",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                if (confirm != MessageBoxResult.Yes) return;
            }

            var scheduledAt = _data.SmtpConfig.EnableScheduledSend ? _data.SmtpConfig.ScheduledSendAt : null;
            if (scheduledAt is DateTime when && when > DateTime.Now)
            {
                var waitMs = (int)Math.Min((when - DateTime.Now).TotalMilliseconds, int.MaxValue);
                StatusMessage = $"Envío programado para {when:dd/MM/yyyy HH:mm}. Esperando…";
                await Task.Delay(waitMs);
            }

            var originalCount = pending.Count;
            pending = ApplySendLimits(pending, "normal");
            if (pending.Count == 0)
            {
                StatusMessage = "No se puede iniciar el envío por límites de envío configurados.";
                return;
            }

            IsBusy = true;
            int ok = 0, fail = 0;
            var failDetails = new List<string>();
            _lastFailedLeadKeys.Clear();

            for (int i = 0; i < pending.Count; i++)
            {
                var row = pending[i];
                StatusMessage = $"Enviando {i + 1} de {pending.Count}: {row.Email}…";

                var (success, error) = await SendLeadAsync(row, forceSend: false);
                if (success) ok++;
                else
                {
                    fail++;
                    var err = string.IsNullOrWhiteSpace(error) ? "Error desconocido al enviar." : error;
                    failDetails.Add($"{row.Email}: {err}");
                    _lastFailedLeadKeys.Add(_data.BuildLeadKey(row.Lead));
#if DEBUG
                    Debug.WriteLine($"[SendSelected] ERROR {row.Email} - {err}");
#endif
                }

                var delayMs = Math.Clamp(_data.SmtpConfig.SendDelayMs, 0, 10000);
                if (i < pending.Count - 1 && delayMs > 0)
                    await Task.Delay(delayMs);
            }

            IsBusy = false;
            UpdateCounts();
            RefreshFilter();

            if (fail > 0)
            {
                var preview = string.Join(" | ", failDetails.Take(3));
                var more    = failDetails.Count > 3 ? $" (+{failDetails.Count - 3} más)" : "";
                StatusMessage = $"Envío completado — ✓ {ok} enviados | ✗ {fail} fallidos. Detalle: {preview}{more}";
            }
            else
            {
                StatusMessage = $"Envío completado — ✓ {ok} enviados  |  ✗ {fail} fallidos";
            }

            if (pending.Count < originalCount)
                StatusMessage += $" | Límite aplicado: {pending.Count}/{originalCount}";
        }

        [RelayCommand]
        private void PrevalidateLeads()
        {
            var candidates = Leads.Where(l => !l.AlreadySent && !l.IsDiscarded).ToList();
            if (candidates.Count == 0)
            {
                StatusMessage = "No hay leads pendientes para prevalidar.";
                return;
            }

            var invalidEmailCount = candidates.Count(l => l.IsInvalidEmail);
            var unconfiguredCourseCount = candidates.Count(l => !IsCourseConfigured(l.Course));
            var duplicateCount = candidates.Count(l => l.IsDuplicate);

            if (invalidEmailCount == 0 && unconfiguredCourseCount == 0 && duplicateCount == 0)
            {
                StatusMessage = $"✓  Prevalidación correcta: {candidates.Count} leads listos para envío.";
                return;
            }

            StatusMessage =
                $"Prevalidación: {candidates.Count} revisados | Emails inválidos: {invalidEmailCount} | Cursos sin configurar: {unconfiguredCourseCount} | Duplicados: {duplicateCount}";
        }

        [RelayCommand]
        private void PrevalidateLegacyLeads()
        {
            var candidates = LegacyLeads.Where(l => !l.IsDiscarded).ToList();
            if (candidates.Count == 0)
            {
                StatusMessage = "No hay leads antiguos para prevalidar.";
                return;
            }

            var invalidEmailCount = candidates.Count(l => l.IsInvalidEmail);
            var missingCampaignCourse = LegacySelectedCampaignCourse == null ? 1 : 0;
            var unconfiguredCampaignCourse =
                LegacySelectedCampaignCourse != null && !IsCourseConfigured(LegacySelectedCampaignCourse) ? 1 : 0;
            var duplicateCount = candidates.Count(l => l.IsDuplicate);

            if (invalidEmailCount == 0 && missingCampaignCourse == 0 && unconfiguredCampaignCourse == 0 && duplicateCount == 0)
            {
                StatusMessage = $"✓  Prevalidación correcta: {candidates.Count} leads antiguos listos para campaña.";
                return;
            }

            StatusMessage =
                $"Prevalidación campaña: {candidates.Count} revisados | Emails inválidos: {invalidEmailCount} | Curso objetivo no seleccionado: {missingCampaignCourse} | Curso objetivo sin configurar: {unconfiguredCampaignCourse} | Duplicados: {duplicateCount}";
        }

        [RelayCommand] private void SelectAll()   { foreach (var l in GetVisibleLeads().Where(x => x.CanSelect)) l.IsSelected = true; }
        [RelayCommand] private void DeselectAll() { foreach (var l in Leads) l.IsSelected = false; }

        [RelayCommand]
        private void ClearSelectedLeads()
        {
            var selectedRows = Leads.Where(x => x.IsSelected).ToList();
            var selected = Leads.Count(l => l.IsSelected);
            foreach (var l in selectedRows)
                l.IsSelected = false;

            if (selectedRows.Count > 0)
                RegisterUndo(
                    "Limpieza de selección en leads",
                    () =>
                    {
                        foreach (var row in selectedRows)
                            row.IsSelected = true;
                        StatusMessage = $"↩  Deshecho: restaurada selección de {selectedRows.Count} leads.";
                    });

            StatusMessage = selected == 0
                ? "No había leads seleccionados para limpiar."
                : $"✓  Selección limpiada: {selected} leads.";
        }

        [RelayCommand]
        private void SelectDuplicateLeads()
        {
            foreach (var l in GetVisibleLeads())
                l.IsSelected = l.CanSelect && l.IsDuplicate;

            var count = Leads.Count(l => l.IsSelected);
            StatusMessage = count == 0
                ? "No hay leads duplicados seleccionables."
                : $"✓  Seleccionados {count} leads duplicados.";
        }

        [RelayCommand]
        private void SelectInvalidEmailLeads()
        {
            foreach (var l in GetVisibleLeads())
                l.IsSelected = l.CanSelect && l.IsInvalidEmail;

            var count = Leads.Count(l => l.IsSelected);
            StatusMessage = count == 0
                ? "No hay leads con email inválido seleccionables."
                : $"✓  Seleccionados {count} leads con email inválido.";
        }

        [RelayCommand]
        private void SelectReadyLeads()
        {
            foreach (var l in GetVisibleLeads())
                l.IsSelected = l.CanSelect
                               && !l.IsDuplicate
                               && !l.IsInvalidEmail
                               && IsCourseConfigured(l.Course);

            var count = Leads.Count(l => l.IsSelected);
            StatusMessage = count == 0
                ? "No hay leads listos para envío automático."
                : $"✓  Seleccionados {count} leads listos para envío.";
        }

        [RelayCommand]
        private void SelectIssueLeads()
        {
            foreach (var l in GetVisibleLeads())
                l.IsSelected = l.CanSelect
                               && (l.IsDuplicate || l.IsInvalidEmail || !IsCourseConfigured(l.Course));

            var count = Leads.Count(l => l.IsSelected);
            StatusMessage = count == 0
                ? "No hay leads con incidencias seleccionables."
                : $"⚠  Seleccionados {count} leads con incidencias.";
        }

        [RelayCommand]
        private void SelectFollowUpToday()
        {
            foreach (var l in GetVisibleLeads())
                l.IsSelected = l.CanSelect && IsFollowUpToday(l.NextContact);

            var count = Leads.Count(l => l.IsSelected);
            StatusMessage = count == 0
                ? "No hay leads con seguimiento para hoy."
                : $"✓  Seleccionados {count} leads con seguimiento de hoy.";
        }

        [RelayCommand]
        private void ExportFilteredLeads()
        {
            var visibleRows = FilteredLeads?.Cast<LeadRow>().ToList() ?? new List<LeadRow>();
            if (!TryGetReportFromDate(out var fromDate, out var dateError))
            {
                StatusMessage = dateError!;
                return;
            }

            var rows = ApplyReportFromDateFilter(visibleRows, fromDate);
            if (rows.Count == 0)
            {
                StatusMessage = "No hay leads filtrados para exportar.";
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "CSV|*.csv",
                FileName = $"leads_{DateTime.Now:yyyyMMdd_HHmm}.csv",
                Title = "Exportar leads filtrados"
            };
            if (dlg.ShowDialog() != true) return;

            var sb = new StringBuilder();
            sb.AppendLine("Estado,Etiqueta,Captadora,Nombre,Email,Curso,Plataforma,Provincia,SituacionLaboral,Telefono,FechaLead,EmailEnviado,WhatsAppEnviado,LlamadoTelefono,FechaLlamadaTelefono,FechaEnvioEmail,FechaEnvioWhatsApp,NotaSeguimiento,ProximoContacto");
            foreach (var r in rows)
            {
                sb.AppendLine(string.Join(",",
                    Csv(r.EstadoTexto), Csv(GetLeadLabelDisplay(r.Label)), Csv(r.HandledBy), Csv(r.Nombre), Csv(r.Email), Csv(r.NombreCurso),
                    Csv(r.Plataforma), Csv(r.Provincia), Csv(r.SituacionLaboral), Csv(r.Telefono),
                    Csv(r.Lead.FechaRegistro),
                    Csv(r.EmailSentAt.HasValue ? "Sí" : "No"),
                    Csv(r.WhatsAppSentAt.HasValue ? "Sí" : "No"),
                    Csv(r.PhoneCalledAt.HasValue ? "Sí" : "No"),
                    Csv(r.PhoneCalledAt?.ToString("dd/MM/yyyy HH:mm") ?? string.Empty),
                    Csv(r.EmailSentAt?.ToString("dd/MM/yyyy HH:mm") ?? string.Empty),
                    Csv(r.WhatsAppSentAt?.ToString("dd/MM/yyyy HH:mm") ?? string.Empty),
                    Csv(r.Note), Csv(r.NextContact)));
            }

            File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
            StatusMessage = $"✓  Exportados {rows.Count} leads a CSV.";
        }

        [RelayCommand]
        private void ExportFilteredLeadsExcel()
        {
            var visibleRows = FilteredLeads?.Cast<LeadRow>().ToList() ?? new List<LeadRow>();
            if (!TryGetReportFromDate(out var fromDate, out var dateError))
            {
                StatusMessage = dateError!;
                return;
            }

            var rows = ApplyReportFromDateFilter(visibleRows, fromDate);
            if (rows.Count == 0)
            {
                StatusMessage = "No hay leads filtrados para exportar.";
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "Excel|*.xlsx",
                FileName = $"leads_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
                Title = "Exportar leads filtrados a Excel"
            };
            if (dlg.ShowDialog() != true) return;
            File.WriteAllBytes(dlg.FileName, BuildLeadsReportExcelBytes(rows));
            StatusMessage = $"✓  Exportados {rows.Count} leads a Excel.";
        }

        public async Task<bool> SendClosingLeadsReportAsync(bool showStatus = false)
        {
            try
            {
                var toEmail = (SmtpConfig.CloseReportRecipientEmail ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(toEmail))
                {
                    if (showStatus) StatusMessage = "Define el email de reporte de cierre en Configuración.";
                    return false;
                }

                if (Leads.Count == 0)
                {
                    if (showStatus) StatusMessage = "No hay leads cargados para generar el reporte.";
                    return false;
                }

                if (!SmtpOkForBackground())
                {
                    if (showStatus) StatusMessage = "Completa la configuración SMTP antes de enviar el reporte.";
                    return false;
                }

                if (!TryGetReportFromDate(out var fromDate, out var dateError))
                {
                    AppLogger.Warn(dateError ?? "Fecha de reporte no válida.");
                    if (showStatus) StatusMessage = dateError ?? "Fecha de reporte no válida.";
                    return false;
                }

                var rows = ApplyReportFromDateFilter(Leads.ToList(), fromDate);
                if (rows.Count == 0)
                {
                    if (showStatus) StatusMessage = "No hay leads en el rango de fecha configurado para el reporte.";
                    return false;
                }

                var now = DateTime.Now;
                var captadora = GetCurrentCaptadoraName();
                var subject = string.IsNullOrWhiteSpace(captadora)
                    ? $"Reporte de Leads {now:dd/MM/yyyy HH:mm}"
                    : $"Reporte de Leads {captadora} {now:dd/MM/yyyy HH:mm}";

                var attachmentName = $"{BuildClosingReportFileNameBase()}.xlsx";
                var attachmentBytes = BuildLeadsReportExcelBytes(rows);
                var body = "Adjunto se envía el reporte automático de leads al cierre de la aplicación.";

                var (success, error) = await _emailService.SendReportAsync(
                    SmtpConfig,
                    toEmail,
                    subject,
                    body,
                    attachmentName,
                    attachmentBytes);

                if (success)
                {
                    AppLogger.Info($"Reporte automático de leads enviado a {toEmail}.");
                    if (showStatus) StatusMessage = $"✓  Reporte de leads enviado a {toEmail}.";
                    return true;
                }
                else
                {
                    AppLogger.Warn($"No se pudo enviar el reporte automático de leads: {error}");
                    if (showStatus) StatusMessage = $"✗  No se pudo enviar el reporte: {error}";
                    return false;
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error("Error al enviar reporte automático de leads al cierre.", ex);
                if (showStatus) StatusMessage = $"✗  Error al enviar reporte: {ex.Message}";
                return false;
            }
        }

        [RelayCommand]
        private async Task SendClosingLeadsReportManual()
        {
            await SendClosingLeadsReportAsync(showStatus: true);
        }

        [RelayCommand]
        private void DownloadClosingLeadsReport()
        {
            if (!TryGetReportFromDate(out var fromDate, out var dateError))
            {
                StatusMessage = dateError!;
                return;
            }

            var rows = ApplyReportFromDateFilter(Leads.ToList(), fromDate);
            if (rows.Count == 0)
            {
                StatusMessage = "No hay leads en el rango de fecha configurado para generar el reporte.";
                return;
            }

            var now = DateTime.Now;
            var dlg = new SaveFileDialog
            {
                Filter = "Excel|*.xlsx",
                FileName = $"{BuildClosingReportFileNameBase()}.xlsx",
                Title = "Guardar reporte de leads"
            };

            if (dlg.ShowDialog() != true) return;

            File.WriteAllBytes(dlg.FileName, BuildLeadsReportExcelBytes(rows));
            StatusMessage = $"✓  Reporte guardado en: {dlg.FileName}";
        }

        private bool TryGetReportFromDate(out DateTime? fromDate, out string? error)
        {
            fromDate = null;
            error = null;

            var raw = (SmtpConfig.ReportFromDateText ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw)) return true;

            var formats = new[]
            {
                "dd/MM/yyyy", "d/M/yyyy", "dd/MM/yy", "d/M/yy",
                "dd-MM-yyyy", "d-M-yyyy", "yyyy-MM-dd",
                "dd/MM/yyyy HH:mm", "d/M/yyyy H:mm"
            };

            if (DateTime.TryParseExact(raw, formats, new CultureInfo("es-ES"), DateTimeStyles.AssumeLocal, out var parsed)
                || DateTime.TryParse(raw, new CultureInfo("es-ES"), DateTimeStyles.AssumeLocal, out parsed)
                || DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out parsed))
            {
                fromDate = parsed.Date;
                return true;
            }

            error = "Formato de 'Fecha desde reporte' no válido. Usa por ejemplo: 25/03/2026";
            return false;
        }

        private List<LeadRow> ApplyReportFromDateFilter(IEnumerable<LeadRow> sourceRows, DateTime? fromDate)
        {
            var rows = sourceRows.ToList();
            if (!fromDate.HasValue) return rows;

            return rows.Where(r =>
            {
                var leadDate = ParseLeadDate(r.Lead.FechaRegistro);
                return leadDate.HasValue && leadDate.Value.Date >= fromDate.Value.Date;
            }).ToList();
        }

        private static DateTime? ParseLeadDate(string? value)
        {
            var raw = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw)) return null;

            var formats = new[]
            {
                "dd/MM/yyyy", "d/M/yyyy", "dd/MM/yy", "d/M/yy",
                "dd-MM-yyyy", "d-M-yyyy", "yyyy-MM-dd",
                "dd/MM/yyyy HH:mm", "d/M/yyyy H:mm", "dd-MM-yyyy HH:mm", "yyyy-MM-dd HH:mm:ss"
            };

            if (DateTime.TryParseExact(raw, formats, new CultureInfo("es-ES"), DateTimeStyles.AssumeLocal, out var dt)
                || DateTime.TryParse(raw, new CultureInfo("es-ES"), DateTimeStyles.AssumeLocal, out dt)
                || DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out dt))
                return dt;

            return null;
        }

        private string BuildClosingReportFileNameBase()
        {
            var captadora = GetCurrentCaptadoraName();
            var namePart = string.IsNullOrWhiteSpace(captadora) ? string.Empty : "_" + captadora.Trim();
            var baseName = $"Leads_Asturias{namePart}";

            foreach (var invalid in Path.GetInvalidFileNameChars())
                baseName = baseName.Replace(invalid, '_');

            return baseName;
        }

        private byte[] BuildLeadsReportExcelBytes(IReadOnlyList<LeadRow> rows)
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Leads");

            var headers = new[]
            {
                "Estado", "Etiqueta", "Captadora", "Nombre", "Email", "Curso", "Plataforma", "Provincia", "SituacionLaboral", "Telefono",
                "FechaLead", "EmailEnviado", "WhatsAppEnviado", "LlamadoTelefono", "FechaLlamadaTelefono", "FechaEnvioEmail", "FechaEnvioWhatsApp",
                "NotaSeguimiento", "ProximoContacto"
            };

            for (int i = 0; i < headers.Length; i++)
                ws.Cells[1, i + 1].Value = headers[i];

            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                var row = i + 2;
                ws.Cells[row, 1].Value = r.EstadoTexto;
                ws.Cells[row, 2].Value = GetLeadLabelDisplay(r.Label);
                ws.Cells[row, 3].Value = r.HandledBy;
                ws.Cells[row, 4].Value = r.Nombre;
                ws.Cells[row, 5].Value = r.Email;
                ws.Cells[row, 6].Value = r.NombreCurso;
                ws.Cells[row, 7].Value = r.Plataforma;
                ws.Cells[row, 8].Value = r.Provincia;
                ws.Cells[row, 9].Value = r.SituacionLaboral;
                ws.Cells[row, 10].Value = r.Telefono;
                ws.Cells[row, 11].Value = r.Lead.FechaRegistro;
                ws.Cells[row, 12].Value = r.EmailSentAt.HasValue ? "Sí" : "No";
                ws.Cells[row, 13].Value = r.WhatsAppSentAt.HasValue ? "Sí" : "No";
                ws.Cells[row, 14].Value = r.PhoneCalledAt.HasValue ? "Sí" : "No";
                ws.Cells[row, 15].Value = r.PhoneCalledAt?.ToString("dd/MM/yyyy HH:mm") ?? string.Empty;
                ws.Cells[row, 16].Value = r.EmailSentAt?.ToString("dd/MM/yyyy HH:mm") ?? string.Empty;
                ws.Cells[row, 17].Value = r.WhatsAppSentAt?.ToString("dd/MM/yyyy HH:mm") ?? string.Empty;
                ws.Cells[row, 18].Value = r.Note;
                ws.Cells[row, 19].Value = r.NextContact;
            }

            if (ws.Dimension != null)
                ws.Cells[ws.Dimension.Address].AutoFitColumns();

            return package.GetAsByteArray();
        }

        private bool SmtpOkForBackground()
            => !string.IsNullOrWhiteSpace(SmtpConfig.Host)
               && SmtpConfig.Port > 0
               && !string.IsNullOrWhiteSpace(SmtpConfig.Username)
               && !string.IsNullOrWhiteSpace(SmtpConfig.Password)
               && (!string.IsNullOrWhiteSpace(SmtpConfig.FromEmail) || !string.IsNullOrWhiteSpace(SmtpConfig.Username));

        [RelayCommand]
        private void ExportLeadIssuesCsv()
        {
            var rows = (FilteredLeads?.Cast<LeadRow>().ToList() ?? new List<LeadRow>())
                .Where(HasLeadIssues)
                .ToList();

            if (rows.Count == 0)
            {
                StatusMessage = "No hay incidencias en los leads filtrados para exportar.";
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "CSV|*.csv",
                FileName = $"leads_incidencias_{DateTime.Now:yyyyMMdd_HHmm}.csv",
                Title = "Exportar incidencias de leads"
            };
            if (dlg.ShowDialog() != true) return;

            var sb = new StringBuilder();
            sb.AppendLine("Nombre,Email,Curso,Incidencias");
            foreach (var r in rows)
            {
                sb.AppendLine(string.Join(",",
                    Csv(r.Nombre),
                    Csv(r.Email),
                    Csv(r.NombreCurso),
                    Csv(BuildLeadIssueText(r))));
            }

            File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
            StatusMessage = $"✓  Exportadas {rows.Count} incidencias de leads.";
        }

        [RelayCommand]
        private void ExportLegacyIssuesCsv()
        {
            var rows = (FilteredLegacyLeads?.Cast<LeadRow>().ToList() ?? new List<LeadRow>())
                .Where(HasLegacyIssues)
                .ToList();

            if (rows.Count == 0)
            {
                StatusMessage = "No hay incidencias en leads antiguos filtrados para exportar.";
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "CSV|*.csv",
                FileName = $"legacy_incidencias_{DateTime.Now:yyyyMMdd_HHmm}.csv",
                Title = "Exportar incidencias de leads antiguos"
            };
            if (dlg.ShowDialog() != true) return;

            var sb = new StringBuilder();
            sb.AppendLine("Nombre,Email,CursoOrigen,CursoCampana,Incidencias");
            foreach (var r in rows)
            {
                sb.AppendLine(string.Join(",",
                    Csv(r.Nombre),
                    Csv(r.Email),
                    Csv(r.Lead.Curso),
                    Csv(LegacySelectedCampaignCourse?.NombreCorto ?? ""),
                    Csv(BuildLegacyIssueText(r))));
            }

            File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
            StatusMessage = $"✓  Exportadas {rows.Count} incidencias de leads antiguos.";
        }

        [RelayCommand]
        private async Task RetryFailed()
        {
            var failedSet = _lastFailedLeadKeys.Count > 0
                ? _lastFailedLeadKeys.ToHashSet(StringComparer.Ordinal)
                : _data.SentRecords
                    .GroupBy(r => r.LeadKey)
                    .Select(g => g.OrderByDescending(x => x.FechaEnvio).First())
                    .Where(r => !r.Success)
                    .Select(r => r.LeadKey)
                    .ToHashSet(StringComparer.Ordinal);

            var retryRows = Leads
                .Where(l => !l.IsDiscarded && !l.AlreadySent && failedSet.Contains(_data.BuildLeadKey(l.Lead)))
                .ToList();

            if (retryRows.Count == 0)
            {
                StatusMessage = "No hay leads fallidos pendientes para reintentar.";
                return;
            }

            foreach (var l in Leads) l.IsSelected = false;
            foreach (var r in retryRows) r.IsSelected = true;

            await SendSelected();
        }

        [RelayCommand]
        private async Task ResendLead(LeadRow? row)
        {
            if (row == null) return;
            if (!SmtpOk()) return;

            IsBusy = true;
            StatusMessage = $"Reenviando a {row.Email}…";
            var (success, error) = await SendLeadAsync(row, forceSend: true);
            IsBusy = false;

            UpdateCounts();
            RefreshFilter();
            StatusMessage = success
                ? $"✓  Reenvío correcto a {row.Email}."
                : $"✗  Error en reenvío a {row.Email}: {error}";
        }

        // ── Gestión de leads ──────────────────────────────────────────────────

        [RelayCommand]
        private void ToggleDiscardLead(LeadRow? row)
        {
            if (row == null) return;
            var previousLabel = row.Label;
            row.Label = row.IsDiscarded ? LeadLabel.Ninguna : LeadLabel.DescartadoNoInteresa;
            RegisterUndo(
                $"Cambio de etiqueta en '{row.Nombre}'",
                () =>
                {
                    row.Label = previousLabel;
                    StatusMessage = $"↩  Deshecho: restaurada etiqueta de '{row.Nombre}'.";
                });
            // persistencia y refresco gestionados por el callback onLabelChanged
        }

        [RelayCommand]
        private void DeleteLead(LeadRow? row)
        {
            if (row == null) return;
            var index = Leads.IndexOf(row);
            var key = _data.BuildLeadKey(row.Lead);
            var previousLabel = _data.GetLeadLabel(key);
            var previousNote = _data.GetLeadNote(key);
            var previousNextContact = _data.GetLeadNextContact(key);

            Leads.Remove(row);
            _data.RemoveLeadLabel(key);
            _data.RemoveLeadNote(key);
            _data.RemoveLeadNextContact(key);

            RegisterUndo(
                $"Eliminación de lead '{row.Nombre}'",
                () =>
                {
                    if (index >= 0 && index <= Leads.Count)
                        Leads.Insert(index, row);
                    else
                        Leads.Add(row);

                    if (previousLabel != LeadLabel.Ninguna) _data.SetLeadLabel(key, previousLabel);
                    if (!string.IsNullOrWhiteSpace(previousNote)) _data.SetLeadNote(key, previousNote);
                    if (!string.IsNullOrWhiteSpace(previousNextContact)) _data.SetLeadNextContact(key, previousNextContact);

                    UpdateCounts();
                    RefreshFilter();
                    StatusMessage = $"↩  Deshecho: lead '{row.Nombre}' restaurado.";
                });

            UpdateCounts();
            RefreshFilter();
            RefreshStudents();
            StatusMessage = $"✓  Lead '{row.Nombre}' eliminado de la lista.";
        }

        // ── Cursos ────────────────────────────────────────────────────────────

        [RelayCommand]
        private void EditCourse(CourseInfo? course)
        {
            if (course == null) return;
            EditingCourse = Clone(course);
            WhatsappMessage = course.TextoWhatsApp ?? "";
            SocialMessage = course.TextoSocial ?? "";
            CourseStartMessage = course.TextoInicioCurso ?? "";
            if (string.IsNullOrWhiteSpace(EditingCourse.TextoMarketingRich) && !string.IsNullOrWhiteSpace(EditingCourse.TextoMarketing))
                EditingCourse.TextoMarketingRich = RichTextSerialization.PlainTextToXaml(EditingCourse.TextoMarketing);
            HasGeminiError = false;
            GeminiError    = "";
            CurrentView    = "CourseEdit";
        }

        [RelayCommand]
        private void SaveCourse()
        {
            if (EditingCourse == null) return;
            if (!string.IsNullOrWhiteSpace(EditingCourse.TextoMarketingRich))
                EditingCourse.TextoMarketing = RichTextSerialization.XamlToPlainText(EditingCourse.TextoMarketingRich);
            EditingCourse.TextoWhatsApp = WhatsappMessage;
            EditingCourse.TextoSocial = SocialMessage;
            EditingCourse.TextoInicioCurso = CourseStartMessage;
            _data.UpdateCourse(EditingCourse);
            RefreshCourses();
            UpdateCourseWarning();
            StatusMessage = "✓  Curso guardado correctamente.";
            CurrentView   = "Courses";
        }

        [RelayCommand]
        private void BrowseCoursePdf()
        {
            if (EditingCourse == null) return;
            var dlg = new OpenFileDialog { Filter = "Archivos PDF|*.pdf", Title = "Seleccionar PDF de Ficha de Inscripción" };
            if (dlg.ShowDialog() != true) return;
            EditingCourse.PdfAdjuntoPath = dlg.FileName;
            OnPropertyChanged(nameof(EditingCourse));
        }

        [RelayCommand]
        private void ClearCoursePdf()
        {
            if (EditingCourse == null) return;
            EditingCourse.PdfAdjuntoPath = "";
            OnPropertyChanged(nameof(EditingCourse));
        }

        [RelayCommand]
        private void DeleteCourse(CourseInfo? course)
        {
            if (course == null) return;
            var confirm = MessageBox.Show(
                $"¿Seguro que quieres eliminar el curso '{course.NombreCorto}'?",
                "Confirmar eliminación", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;

            var deleted = _data.DeleteCourse(course.Id);
            if (!deleted) { StatusMessage = "No se pudo eliminar el curso."; return; }

            if (EditingCourse?.Id == course.Id) EditingCourse = null;

            var toRemove = Leads.Where(l => l.Course.Id == course.Id).ToList();
            foreach (var lead in toRemove) Leads.Remove(lead);

            RefreshCourses();
            UpdateCounts();
            RefreshFilter();
            CurrentView   = "Courses";
            StatusMessage = "✓  Curso eliminado correctamente.";
        }

        [RelayCommand]
        private void DeleteUnconfiguredCourses()
        {
            var targets = Courses.Where(c => !IsCourseConfigured(c)).ToList();
            if (targets.Count == 0)
            {
                StatusMessage = "No hay cursos sin configurar para eliminar.";
                return;
            }

            var confirm = MessageBox.Show(
                $"Se eliminarán {targets.Count} curso(s) sin configurar. ¿Deseas continuar?",
                "Eliminar cursos sin configurar",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;

            var deletedIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var course in targets)
            {
                if (_data.DeleteCourse(course.Id))
                    deletedIds.Add(course.Id);
            }

            if (deletedIds.Count == 0)
            {
                StatusMessage = "No se pudo eliminar ningún curso sin configurar.";
                return;
            }

            if (EditingCourse != null && deletedIds.Contains(EditingCourse.Id))
                EditingCourse = null;

            foreach (var lead in Leads.Where(l => deletedIds.Contains(l.Course.Id)).ToList())
                Leads.Remove(lead);

            foreach (var legacy in LegacyLeads.Where(l => deletedIds.Contains(l.Course.Id)).ToList())
                LegacyLeads.Remove(legacy);

            RefreshCourses();
            UpdateCourseWarning();
            UpdateCounts();
            RefreshFilter();
            RefreshLegacyFilter();
            CurrentView = "Courses";
            StatusMessage = $"✓  Eliminados {deletedIds.Count} curso(s) sin configurar.";
        }

        [RelayCommand] private void CancelEdit() { EditingCourse = null; WhatsappMessage = ""; SocialMessage = ""; CurrentView = "Courses"; }

        // ── IA ────────────────────────────────────────────────────────────────
        [RelayCommand]
        private async Task GenerateMarketingText()
        {
            if (EditingCourse == null) return;

            var provider = _data.SmtpConfig.AiProvider == "Gemini" ? AiProvider.Gemini : AiProvider.OpenRouter;
            var apiKey   = provider == AiProvider.Gemini ? _data.SmtpConfig.GeminiApiKey : _data.SmtpConfig.OpenRouterApiKey;

            if (string.IsNullOrWhiteSpace(apiKey))
            {
                GeminiError    = provider == AiProvider.Gemini
                    ? "⚠  Introduce tu API Key de Gemini en Configuración."
                    : "⚠  Introduce tu API Key de OpenRouter en Configuración.";
                HasGeminiError = true;
                return;
            }

            IsGenerating = true; HasGeminiError = false; GeminiError = "";

            var (success, result) = await _aiText.GenerateMarketingTextAsync(
                apiKey:        apiKey,
                provider:      provider,
                nombreCurso:   EditingCourse.NombreCorto,
                requisitos:    EditingCourse.RequisitosAcceso,
                fechaInicio:   EditingCourse.FechaInicio,
                fechaFin:      EditingCourse.FechaFin,
                horario:       EditingCourse.HorarioInfo,
                infoAdicional: EditingCourse.InfoAdicional);

            IsGenerating = false;
            if (success)
            {
                EditingCourse.TextoMarketing     = result;
                EditingCourse.TextoMarketingRich = RichTextSerialization.PlainTextToXaml(result);
                OnPropertyChanged(nameof(EditingCourse));
            }
            else { GeminiError = result; HasGeminiError = true; }
        }

        // ── SMTP ──────────────────────────────────────────────────────────────
        [RelayCommand]
        private void SaveSmtp()
        {
            _data.SmtpConfig.Host             = SmtpConfig.Host;
            _data.SmtpConfig.Port             = SmtpConfig.Port;
            _data.SmtpConfig.EnableSsl        = SmtpConfig.EnableSsl;
            _data.SmtpConfig.Username         = SmtpConfig.Username;
            _data.SmtpConfig.Password         = SmtpConfig.Password;
            _data.SmtpConfig.FromName         = SmtpConfig.FromName;
            _data.SmtpConfig.FromEmail        = SmtpConfig.FromEmail;
            _data.SmtpConfig.GeminiApiKey     = SmtpConfig.GeminiApiKey;
            _data.SmtpConfig.OpenRouterApiKey = SmtpConfig.OpenRouterApiKey;
            _data.SmtpConfig.AiProvider       = SmtpConfig.AiProvider;
            _data.SmtpConfig.SendDelayMs      = Math.Clamp(SmtpConfig.SendDelayMs, 0, 10000);
            _data.SmtpConfig.WhatsAppNumber   = SmtpConfig.WhatsAppNumber;
            _data.SmtpConfig.ConfirmMassSend  = SmtpConfig.ConfirmMassSend;
            _data.SmtpConfig.EnableScheduledSend = SmtpConfig.EnableScheduledSend;
            _data.SmtpConfig.ScheduledSendAt = SmtpConfig.ScheduledSendAt;
            _data.SmtpConfig.MaxSendsPerSession = Math.Max(0, SmtpConfig.MaxSendsPerSession);
            _data.SmtpConfig.MaxSendsPerDay = Math.Max(0, SmtpConfig.MaxSendsPerDay);
            _data.SmtpConfig.LabelTemplateInteresado = SmtpConfig.LabelTemplateInteresado;
            _data.SmtpConfig.LabelTemplateConfirmado = SmtpConfig.LabelTemplateConfirmado;
            _data.SmtpConfig.LabelTemplateDescartado = SmtpConfig.LabelTemplateDescartado;
            _data.SmtpConfig.EmailFooter = SmtpConfig.EmailFooter;
            _data.Save();
            SmtpTestResult = "✓  Configuración guardada.";
        }

        [RelayCommand]
        private void ExportSmtpConfig()
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Configuración LeadMailer|*.json|JSON|*.json|Todos|*.*",
                FileName = $"LeadMailer_config_{DateTime.Now:yyyyMMdd_HHmm}.json",
                Title = "Exportar configuración"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                _data.ExportSmtpConfig(dlg.FileName);
                StatusMessage = "✓  Configuración exportada correctamente.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"✗  Error al exportar configuración: {ex.Message}";
            }
        }

        [RelayCommand]
        private void ImportSmtpConfig()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Configuración LeadMailer|*.json|JSON|*.json|Todos|*.*",
                Title = "Importar configuración"
            };
            if (dlg.ShowDialog() != true) return;

            var confirm = MessageBox.Show(
                "Esto reemplazará los parámetros actuales de configuración. ¿Deseas continuar?",
                "Confirmar importación",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;

            try
            {
                _data.ImportSmtpConfig(dlg.FileName);
                SmtpConfig = _data.SmtpConfig;
                ExcelFilePath = _data.SmtpConfig.ExcelFilePath;
                StatusMessage = "✓  Configuración importada correctamente.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"✗  Error al importar configuración: {ex.Message}";
            }
        }

        [RelayCommand]
        private void ExportAppData()
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Backup LeadMailer|*.json|JSON|*.json|Todos|*.*",
                FileName = $"LeadMailer_backup_{DateTime.Now:yyyyMMdd_HHmm}.json",
                Title = "Exportar copia de datos"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                _data.ExportData(dlg.FileName);
                StatusMessage = "✓  Copia exportada correctamente.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"✗  Error al exportar copia: {ex.Message}";
            }
        }

        [RelayCommand]
        private void ImportAppData()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Backup LeadMailer|*.json|JSON|*.json|Todos|*.*",
                Title = "Importar copia de datos"
            };
            if (dlg.ShowDialog() != true) return;

            var confirm = MessageBox.Show(
                "Esto reemplazará los datos actuales (historial, cursos y configuración). ¿Deseas continuar?",
                "Confirmar importación",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;

            try
            {
                _data.ImportData(dlg.FileName);

                SmtpConfig = _data.SmtpConfig;
                ExcelFilePath = _data.SmtpConfig.ExcelFilePath;

                foreach (var row in LegacyLeads)
                    row.PropertyChanged -= OnLegacyLeadPropertyChanged;

                Leads.Clear();
                LegacyLeads.Clear();

                RefreshCourses();
                RefreshHistory();
                RefreshLegacyFilterOptions();
                RefreshLegacyFilter();
                UpdateCounts();
                RefreshFilter();
                RefreshStudents();
                UpdateCourseWarning();

                StatusMessage = "✓  Copia importada correctamente.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"✗  Error al importar copia: {ex.Message}";
            }
        }

        [RelayCommand]
        private async Task TestSmtp()
        {
            if (!SmtpOk()) return;
            IsBusy = true; SmtpTestResult = "Probando conexión…";
            var (ok, err) = await _emailService.TestConnectionAsync(SmtpConfig);
            SmtpTestResult = ok ? "✓  Conexión correcta. El servidor SMTP responde." : $"✗  Error: {err}";
            IsBusy = false;
        }

        [RelayCommand]
        private void ClearHistory()
        {
            var confirm = MessageBox.Show(
                "¿Seguro que quieres borrar el historial de envíos?",
                "Confirmar borrado", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;

            _data.SentRecords.Clear();
            _data.Save();
            RefreshHistory();
            StatusMessage = "✓  Historial de envíos eliminado.";
        }

        [RelayCommand]
        private void ExportHistoryCsv()
        {
            var rows = FilteredHistory?.Cast<SentRecord>().ToList() ?? new List<SentRecord>();
            if (rows.Count == 0)
            {
                StatusMessage = "No hay registros en el historial para exportar.";
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "CSV|*.csv",
                FileName = $"historial_envios_{DateTime.Now:yyyyMMdd_HHmm}.csv",
                Title = "Exportar historial de envíos"
            };
            if (dlg.ShowDialog() != true) return;

            var sb = new StringBuilder();
            sb.AppendLine("Fecha,Resultado,Nombre,Email,Curso,Error");
            foreach (var r in rows)
            {
                sb.AppendLine(string.Join(",",
                    Csv(r.FechaEnvio.ToString("dd/MM/yyyy HH:mm")),
                    Csv(r.Success ? "Correcto" : "Fallido"),
                    Csv(r.NombreLead ?? string.Empty),
                    Csv(r.Email),
                    Csv(r.CursoRaw),
                    Csv(r.Error ?? string.Empty)));
            }

            File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
            StatusMessage = $"✓  Exportados {rows.Count} registros del historial.";
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private void BuildFilteredView()
        {
            _filteredLeads = CollectionViewSource.GetDefaultView(Leads);
            _filteredLeads.Filter = LeadsFilter;
            _filteredLeads.SortDescriptions.Add(new SortDescription(nameof(LeadRow.NombreCurso), ListSortDirection.Ascending));
            _filteredLeads.SortDescriptions.Add(new SortDescription(nameof(LeadRow.Nombre),      ListSortDirection.Ascending));
            FilteredCount = 0;
        }

        private void BuildHistoryView()
        {
            _filteredHistory = CollectionViewSource.GetDefaultView(SentHistory);
            _filteredHistory.Filter = HistoryFilter;
            _filteredHistory.SortDescriptions.Add(new SortDescription(nameof(SentRecord.FechaEnvio), ListSortDirection.Descending));
        }

        private IEnumerable<LeadRow> GetVisibleLeads()
            => (FilteredLeads?.Cast<LeadRow>().ToList() ?? Leads.ToList());

        private void BuildStudentsView()
        {
            _filteredStudents = CollectionViewSource.GetDefaultView(Students);
            _filteredStudents.Filter = StudentsFilter;
            _filteredStudents.SortDescriptions.Add(new SortDescription(nameof(StudentEntry.CursoDisplay), ListSortDirection.Ascending));
            _filteredStudents.SortDescriptions.Add(new SortDescription(nameof(StudentEntry.Nombre), ListSortDirection.Ascending));
            _filteredStudents.GroupDescriptions.Add(new PropertyGroupDescription(nameof(StudentEntry.CursoDisplay)));
        }

        private void BuildLegacyFilteredView()
        {
            LegacyOriginCourseFilterOptions.Clear();
            LegacyOriginCourseFilterOptions.Add("Todos");
            LegacyProvinceFilterOptions.Clear();
            LegacyProvinceFilterOptions.Add("Todas");
            LegacyEmploymentFilterOptions.Clear();
            LegacyEmploymentFilterOptions.Add("Todas");

            _filteredLegacyLeads = CollectionViewSource.GetDefaultView(LegacyLeads);
            _filteredLegacyLeads.Filter = LegacyLeadsFilter;
            _filteredLegacyLeads.SortDescriptions.Add(new SortDescription(nameof(LeadRow.Nombre), ListSortDirection.Ascending));
            LegacyFilteredCount = 0;
            RefreshLegacyIssuesSummary();
        }

        private bool LeadsFilter(object obj)
        {
            if (obj is not LeadRow row) return false;

            // Filtro por curso
            if (!string.IsNullOrEmpty(SelectedCourseFilter?.Id) && row.Course.Id != SelectedCourseFilter.Id)
                return false;

            // Filtro por estado
            if (SelectedStatusFilter != "Todos")
            {
                var match = SelectedStatusFilter switch
                {
                    "Pendiente"        => row.Status == LeadStatus.Pendiente,
                    "Email Enviado"    => row.Status == LeadStatus.EmailEnviado,
                    "WhatsApp Enviado" => row.Status == LeadStatus.WhatsAppEnviado,
                    "Ambos Enviados"   => row.Status == LeadStatus.AmbosEnviados,
                    "Incidencias"      => row.IsDuplicate || row.IsInvalidEmail || !IsCourseConfigured(row.Course),
                    "Seguimiento hoy"  => IsFollowUpToday(row.NextContact),
                    _                   => true
                };
                if (!match) return false;
            }

            // Filtro por texto libre
            if (!string.IsNullOrWhiteSpace(SearchText))
            {
                var q = SearchText.Trim().ToLowerInvariant();
                if (!row.Nombre.ToLowerInvariant().Contains(q)
                 && !row.Email.ToLowerInvariant().Contains(q)
                 && !row.Telefono.ToLowerInvariant().Contains(q)
                 && !row.Provincia.ToLowerInvariant().Contains(q)
                 && !row.NombreCurso.ToLowerInvariant().Contains(q)
                 && !row.Note.ToLowerInvariant().Contains(q)
                 && !row.NextContact.ToLowerInvariant().Contains(q))
                    return false;
            }

            return true;
        }

        private bool StudentsFilter(object obj)
        {
            if (obj is not StudentEntry row) return false;

            if (!string.IsNullOrWhiteSpace(SelectedStudentCourseFilter)
                && SelectedStudentCourseFilter != "Todos"
                && NormalizeStudentCourseKey(row.CursoDisplay) != NormalizeStudentCourseKey(SelectedStudentCourseFilter)
                && NormalizeStudentCourseKey(row.CursoRaw) != NormalizeStudentCourseKey(SelectedStudentCourseFilter))
                return false;

            if (!string.IsNullOrWhiteSpace(StudentSearchText))
            {
                var q = StudentSearchText.Trim().ToLowerInvariant();
                if (!row.Nombre.ToLowerInvariant().Contains(q)
                 && !row.Email.ToLowerInvariant().Contains(q)
                 && !row.CursoDisplay.ToLowerInvariant().Contains(q))
                    return false;
            }

            return true;
        }

        private static string NormalizeStudentCourseKey(string? value)
            => (value ?? string.Empty).Trim().ToLowerInvariant();

        private bool HasLeadIssues(LeadRow row)
            => row.IsDuplicate || row.IsInvalidEmail || !IsCourseConfigured(row.Course);

        private string BuildLeadIssueText(LeadRow row)
        {
            var issues = new List<string>();
            if (row.IsDuplicate) issues.Add("Duplicado");
            if (row.IsInvalidEmail) issues.Add("Email inválido");
            if (!IsCourseConfigured(row.Course)) issues.Add("Curso sin configurar");
            return string.Join(" | ", issues);
        }

        private bool HasLegacyIssues(LeadRow row)
        {
            var campaignIssue = LegacySelectedCampaignCourse == null || !IsCourseConfigured(LegacySelectedCampaignCourse);
            return row.IsDuplicate || row.IsInvalidEmail || campaignIssue;
        }

        private string BuildLegacyIssueText(LeadRow row)
        {
            var issues = new List<string>();
            if (row.IsDuplicate) issues.Add("Duplicado");
            if (row.IsInvalidEmail) issues.Add("Email inválido");
            if (LegacySelectedCampaignCourse == null) issues.Add("Curso campaña no seleccionado");
            else if (!IsCourseConfigured(LegacySelectedCampaignCourse)) issues.Add("Curso campaña sin configurar");
            return string.Join(" | ", issues);
        }

        private bool LegacyLeadsFilter(object obj)
        {
            if (obj is not LeadRow row) return false;

            if (LegacyOriginCourseFilter != "Todos"
                && !string.Equals(row.Lead.Curso, LegacyOriginCourseFilter, StringComparison.OrdinalIgnoreCase))
                return false;

            if (LegacyProvinceFilter != "Todas"
                && !string.Equals(row.Provincia, LegacyProvinceFilter, StringComparison.OrdinalIgnoreCase))
                return false;

            if (LegacyEmploymentFilter != "Todas"
                && !string.Equals(row.SituacionLaboral, LegacyEmploymentFilter, StringComparison.OrdinalIgnoreCase))
                return false;

            if (LegacyOnlyWithPublicidad && !row.Lead.AceptaPublicidad)
                return false;

            if (LegacyOnlyDuplicates && !row.IsDuplicate)
                return false;

            if (LegacyOnlyInvalidEmail && !row.IsInvalidEmail)
                return false;

            if (LegacyOnlyIssues)
            {
                var campaignHasIssues = LegacySelectedCampaignCourse == null || !IsCourseConfigured(LegacySelectedCampaignCourse);
                var rowHasIssues = row.IsDuplicate || row.IsInvalidEmail || campaignHasIssues;
                if (!rowHasIssues) return false;
            }

            if (!string.IsNullOrWhiteSpace(LegacySearchText))
            {
                var q = LegacySearchText.Trim().ToLowerInvariant();
                if (!row.Nombre.ToLowerInvariant().Contains(q)
                 && !row.Email.ToLowerInvariant().Contains(q)
                 && !row.Telefono.ToLowerInvariant().Contains(q)
                 && !row.Provincia.ToLowerInvariant().Contains(q)
                 && !row.SituacionLaboral.ToLowerInvariant().Contains(q)
                 && !row.Lead.Curso.ToLowerInvariant().Contains(q))
                    return false;
            }

            return true;
        }

        private static bool IsFollowUpToday(string? nextContact)
        {
            if (string.IsNullOrWhiteSpace(nextContact)) return false;

            var value = nextContact.Trim().ToLowerInvariant();
            if (value.Contains("hoy")) return true;

            var today = DateTime.Today;
            var todayTokens = new[]
            {
                today.ToString("d/M", CultureInfo.InvariantCulture),
                today.ToString("dd/MM", CultureInfo.InvariantCulture),
                today.ToString("d-M", CultureInfo.InvariantCulture),
                today.ToString("dd-MM", CultureInfo.InvariantCulture),
            };

            return todayTokens.Any(t => value.Contains(t));
        }

        private bool HistoryFilter(object obj)
        {
            if (obj is not SentRecord row) return false;

            if (HistoryFromDate is DateTime fromDate && row.FechaEnvio.Date < fromDate.Date)
                return false;

            if (HistoryToDate is DateTime toDate && row.FechaEnvio.Date > toDate.Date)
                return false;

            if (HistoryStatusFilter != "Todos")
            {
                var match = HistoryStatusFilter switch
                {
                    "Correctos" => row.Success,
                    "Fallidos"  => !row.Success,
                    _            => true
                };
                if (!match) return false;
            }

            if (!string.IsNullOrWhiteSpace(HistorySearchText))
            {
                var q = HistorySearchText.Trim().ToLowerInvariant();
                var nombre = (row.NombreLead ?? "").ToLowerInvariant();
                if (!nombre.Contains(q)
                 && !row.Email.ToLowerInvariant().Contains(q)
                 && !row.CursoRaw.ToLowerInvariant().Contains(q)
                 && !(row.Error ?? "").ToLowerInvariant().Contains(q))
                    return false;
            }

            return true;
        }

        private void RefreshFilter()
        {
            if (_filteredLeads is System.ComponentModel.IEditableCollectionView ecv)
            {
                if (ecv.IsEditingItem) ecv.CommitEdit();
                if (ecv.IsAddingNew)   ecv.CommitNew();
            }
            _filteredLeads?.Refresh();
            FilteredCount = _filteredLeads?.Cast<LeadRow>().Count() ?? 0;
            RefreshLeadIssuesSummary();
        }

        private void RefreshLeadIssuesSummary()
        {
            var rows = (FilteredLeads?.Cast<LeadRow>().ToList() ?? new List<LeadRow>());
            LeadIssuesDuplicateCount = rows.Count(r => r.IsDuplicate);
            LeadIssuesInvalidEmailCount = rows.Count(r => r.IsInvalidEmail);
            LeadIssuesUnconfiguredCourseCount = rows.Count(r => !IsCourseConfigured(r.Course));
            LeadIssuesTotalCount = rows.Count(r => HasLeadIssues(r));
        }

        private void RefreshHistoryFilter()
        {
            _filteredHistory?.Refresh();
        }

        private void RefreshStudentsFilter()
        {
            _filteredStudents?.Refresh();
        }

        private void RefreshLegacyFilter()
        {
            _filteredLegacyLeads?.Refresh();
            LegacyFilteredCount = _filteredLegacyLeads?.Cast<LeadRow>().Count() ?? 0;
            RefreshLegacySelectionCount();
            RefreshLegacyIssuesSummary();
        }

        private void RefreshStudents()
        {
            Students.Clear();
            StudentCourseFilterOptions.Clear();
            StudentCourseFilterOptions.Add("Todos");

            foreach (var course in Courses.Where(IsCourseConfigured).OrderBy(c => c.NombreCorto))
                AddStudentCourseFilterOption(course.NombreCorto);

            foreach (var lead in Leads.Where(l => l.Label == LeadLabel.ConfirmaSi))
            {
                var row = new StudentEntry(lead);
                RegisterStudentRow(row);
                Students.Add(row);
            }

            foreach (var student in _data.Students)
            {
                var row = new StudentEntry(student, _data.Save);
                RegisterStudentRow(row);
                Students.Add(row);
            }

            if (!StudentCourseFilterOptions.Contains(SelectedStudentCourseFilter))
                SelectedStudentCourseFilter = "Todos";

            RefreshStudentsFilter();
        }

        private void AddStudentCourseFilterOption(string? course)
        {
            if (string.IsNullOrWhiteSpace(course)) return;
            var trimmed = course.Trim();
            if (!StudentCourseFilterOptions.Any(x => string.Equals(x, trimmed, StringComparison.OrdinalIgnoreCase)))
                StudentCourseFilterOptions.Add(trimmed);
        }

        private void RefreshLegacyIssuesSummary()
        {
            var rows = (FilteredLegacyLeads?.Cast<LeadRow>().ToList() ?? new List<LeadRow>());
            LegacyIssuesDuplicateCount = rows.Count(r => r.IsDuplicate);
            LegacyIssuesInvalidEmailCount = rows.Count(r => r.IsInvalidEmail);

            var campaignIssue = LegacySelectedCampaignCourse == null || !IsCourseConfigured(LegacySelectedCampaignCourse);
            LegacyIssuesCampaignCourseCount = campaignIssue ? rows.Count : 0;
            LegacyIssuesTotalCount = rows.Count(HasLegacyIssues);
        }

        private void RefreshLegacySentStateForSelectedCampaignCourse()
        {
            var selectedCourseRaw = LegacySelectedCampaignCourse?.CursoRaw;

            foreach (var row in LegacyLeads)
            {
                var key = _data.BuildLeadKey(row.Lead);
                var emailSentAt = _data.GetLastSuccessfulEmailSentAt(key, selectedCourseRaw);
                row.EmailSentAt = emailSentAt;
                row.AlreadySent = emailSentAt.HasValue;

                if (emailSentAt.HasValue && row.IsSelected)
                    row.IsSelected = false;
            }
        }

        private string BuildSocialText()
        {
            if (!string.IsNullOrWhiteSpace(SocialMessage))
                return SocialMessage.Trim();

            if (EditingCourse == null)
                return string.Empty;

            var fallback = !string.IsNullOrWhiteSpace(EditingCourse.TextoMarketing)
                ? EditingCourse.TextoMarketing
                : RichTextSerialization.XamlToPlainText(EditingCourse.TextoMarketingRich);

            return fallback?.Trim() ?? string.Empty;
        }

        [RelayCommand]
        private void UseMarketingAsSocial()
        {
            if (EditingCourse == null) return;
            var text = !string.IsNullOrWhiteSpace(EditingCourse.TextoMarketing)
                ? EditingCourse.TextoMarketing
                : RichTextSerialization.XamlToPlainText(EditingCourse.TextoMarketingRich);

            if (string.IsNullOrWhiteSpace(text))
            {
                StatusMessage = "No hay texto de marketing disponible para usar en redes sociales.";
                return;
            }

            SocialMessage = text.Trim();
            StatusMessage = "✓  Texto social actualizado desde marketing.";
        }

        [RelayCommand]
        private void CopySocialText()
        {
            var text = BuildSocialText();
            if (string.IsNullOrWhiteSpace(text))
            {
                StatusMessage = "No hay texto social para copiar.";
                return;
            }

            try
            {
                Clipboard.SetText(text);
                StatusMessage = "✓  Texto social copiado al portapapeles.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"✗  No se pudo copiar el texto social: {ex.Message}";
            }
        }

        [RelayCommand]
        private void OpenLinkedIn()
        {
            var text = BuildSocialText();
            if (string.IsNullOrWhiteSpace(text))
            {
                StatusMessage = "No hay texto social para publicar.";
                return;
            }

            try
            {
                Clipboard.SetText(text);
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://www.linkedin.com/feed/",
                    UseShellExecute = true
                });
                StatusMessage = "LinkedIn abierto. Texto copiado al portapapeles.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"✗  No se pudo abrir LinkedIn: {ex.Message}";
            }
        }

        [RelayCommand]
        private void OpenFacebook()
        {
            var text = BuildSocialText();
            if (string.IsNullOrWhiteSpace(text))
            {
                StatusMessage = "No hay texto social para publicar.";
                return;
            }

            try
            {
                Clipboard.SetText(text);
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://www.facebook.com/",
                    UseShellExecute = true
                });
                StatusMessage = "Facebook abierto. Texto copiado al portapapeles.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"✗  No se pudo abrir Facebook: {ex.Message}";
            }
        }

        private void MarkDuplicatesForRows(IEnumerable<LeadRow> rows)
        {
            var list = rows.ToList();
            foreach (var row in list)
            {
                row.IsDuplicate = false;
                row.IsInvalidEmail = !IsValidEmail(row.Email);
            }

            static string N(string? value) => (value ?? string.Empty).Trim().ToLowerInvariant();

            var emailAndCourseGroups = list
                .Where(r => !string.IsNullOrWhiteSpace(r.Email) && !string.IsNullOrWhiteSpace(r.Lead.Curso))
                .GroupBy(r => (Email: N(r.Email), Curso: N(r.Lead.Curso)))
                .Where(g => g.Count() > 1);

            foreach (var group in emailAndCourseGroups)
                foreach (var row in group)
                    row.IsDuplicate = true;
        }

        private static string NormalizePhoneForDuplicate(string? phone)
            => new string((phone ?? string.Empty).Where(char.IsDigit).ToArray());

        private void RefreshHistory()
        {
            SentHistory.Clear();
            var ordered = _data.SentRecords.OrderByDescending(x => x.FechaEnvio).ToList();
            foreach (var r in ordered)
                SentHistory.Add(r);

            HistorySuccessCount = ordered.Count(x => x.Success);
            HistoryFailedCount = ordered.Count(x => !x.Success);

            var today = DateTime.Today;
            HistoryTodayCount = ordered.Count(x => x.FechaEnvio.Date == today);
            HistoryLast7DaysCount = ordered.Count(x => x.FechaEnvio.Date >= today.AddDays(-6));
            HistoryLast30DaysCount = ordered.Count(x => x.FechaEnvio.Date >= today.AddDays(-29));

            HistoryCourseMetrics.Clear();
            foreach (var group in ordered
                         .GroupBy(x => string.IsNullOrWhiteSpace(x.CursoRaw) ? "(Sin curso)" : x.CursoRaw.Trim())
                         .Select(g => new CourseHistoryMetric(
                             Curso: g.Key,
                             Total: g.Count(),
                             Correctos: g.Count(x => x.Success),
                             Fallidos: g.Count(x => !x.Success),
                             SuccessRate: g.Count() == 0 ? 0 : (100.0 * g.Count(x => x.Success) / g.Count())))
                         .OrderByDescending(x => x.Total)
                         .ThenBy(x => x.Curso)
                         .Take(5))
            {
                HistoryCourseMetrics.Add(group);
            }

            RefreshHistoryFilter();
        }

        private void RefreshCourses()
        {
            Courses.Clear();
            var previousId = SelectedCourseFilter?.Id;
            var previousLegacyCourseId = LegacySelectedCampaignCourse?.Id;

            CourseFilterItems.Clear();
            CourseFilterItems.Add(_todosCursos);
            foreach (var c in _data.Courses.OrderBy(x => x.NombreCorto))
            {
                Courses.Add(c);
                CourseFilterItems.Add(c);
            }

            SelectedCourseFilter = string.IsNullOrEmpty(previousId)
                ? _todosCursos
                : CourseFilterItems.FirstOrDefault(c => c.Id == previousId) ?? _todosCursos;

            LegacySelectedCampaignCourse = string.IsNullOrWhiteSpace(previousLegacyCourseId)
                ? null
                : Courses.FirstOrDefault(c => c.Id == previousLegacyCourseId);
        }

        private void RefreshLegacyFilterOptions()
        {
            var selectedLegacyCourse = LegacyOriginCourseFilter;
            var selectedProvince = LegacyProvinceFilter;
            var selectedEmployment = LegacyEmploymentFilter;

            LegacyOriginCourseFilterOptions.Clear();
            LegacyOriginCourseFilterOptions.Add("Todos");
            foreach (var legacyCourse in LegacyLeads
                         .Select(x => x.Lead.Curso?.Trim())
                         .Where(x => !string.IsNullOrWhiteSpace(x))
                         .Distinct(StringComparer.OrdinalIgnoreCase)
                         .OrderBy(x => x))
            {
                LegacyOriginCourseFilterOptions.Add(legacyCourse!);
            }

            LegacyProvinceFilterOptions.Clear();
            LegacyProvinceFilterOptions.Add("Todas");
            foreach (var province in LegacyLeads
                         .Select(x => x.Provincia?.Trim())
                         .Where(x => !string.IsNullOrWhiteSpace(x))
                         .Distinct(StringComparer.OrdinalIgnoreCase)
                         .OrderBy(x => x))
            {
                LegacyProvinceFilterOptions.Add(province!);
            }

            LegacyEmploymentFilterOptions.Clear();
            LegacyEmploymentFilterOptions.Add("Todas");
            foreach (var employment in LegacyLeads
                         .Select(x => x.SituacionLaboral?.Trim())
                         .Where(x => !string.IsNullOrWhiteSpace(x))
                         .Distinct(StringComparer.OrdinalIgnoreCase)
                         .OrderBy(x => x))
            {
                LegacyEmploymentFilterOptions.Add(employment!);
            }

            LegacyOriginCourseFilter = LegacyOriginCourseFilterOptions.Contains(selectedLegacyCourse) ? selectedLegacyCourse : "Todos";
            LegacyProvinceFilter = LegacyProvinceFilterOptions.Contains(selectedProvince) ? selectedProvince : "Todas";
            LegacyEmploymentFilter = LegacyEmploymentFilterOptions.Contains(selectedEmployment) ? selectedEmployment : "Todas";
        }

        private void OnLegacyLeadPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(LeadRow.IsSelected))
                RefreshLegacySelectionCount();
        }

        private void RefreshLegacySelectionCount()
            => LegacySelectedCount = LegacyLeads.Count(x => x.IsSelected);

        private void ClearNewStudentFields()
        {
            NewStudentNombre = "";
            NewStudentEmail = "";
            NewStudentTelefono = "";
            NewStudentCurso = "";
            NewStudentPlataforma = "";
            NewStudentSituacionLaboral = "";
            NewStudentNivelEstudios = "";
            NewStudentProvincia = "";
            NewStudentObservaciones = "";
            NewStudentContacto = "";
            NewStudentInscripcion = "";
            NewStudentFechaRegistro = "";
            NewStudentAceptaPublicidad = true;
        }

        private void UpdateCounts()
        {
            PendingCount      = Leads.Count(l => l.Status == LeadStatus.Pendiente);
            SentCount         = Leads.Count(l => l.EmailSentAt.HasValue);
            FollowUpTodayCount = Leads.Count(l => IsFollowUpToday(l.NextContact));
        }

        private static bool IsCourseConfigured(CourseInfo course)
            => !string.IsNullOrWhiteSpace(course.TextoMarketing)
            || !string.IsNullOrWhiteSpace(course.TextoMarketingRich);

        private void UpdateCourseWarning()
        {
            var unconfigured = Leads
                .Select(l => l.Course)
                .DistinctBy(c => c.Id)
                .Where(c => !IsCourseConfigured(c))
                .Select(c => c.NombreCorto)
                .ToList();

            HasCourseWarning  = unconfigured.Count > 0;
            CourseWarningText = unconfigured.Count > 0
                ? $"Se detectaron {unconfigured.Count} curso(s) sin configurar. Ve a Cursos para completar su información."
                : "";
        }

        private bool SmtpOk()
        {
            if (!string.IsNullOrWhiteSpace(_data.SmtpConfig.Host) && !string.IsNullOrWhiteSpace(_data.SmtpConfig.Username)) return true;
            StatusMessage = "⚠  Configura primero el servidor SMTP (pestaña Configuración).";
            return false;
        }

        [RelayCommand]
        private void UndoLastAction()
        {
            if (_undoAction == null)
            {
                StatusMessage = "No hay acciones para deshacer.";
                return;
            }

            var action = _undoAction;
            _undoAction = null;
            var description = _undoDescription;
            _undoDescription = string.Empty;

            action();
            if (string.IsNullOrWhiteSpace(StatusMessage))
                StatusMessage = $"↩  Deshecho: {description}.";
        }

        private void RegisterUndo(string description, Action action)
        {
            _undoDescription = description;
            _undoAction = action;
        }

        private List<LeadRow> ApplySendLimits(List<LeadRow> rows, string scope)
        {
            var result = rows;

            if (_data.SmtpConfig.MaxSendsPerSession > 0)
                result = result.Take(_data.SmtpConfig.MaxSendsPerSession).ToList();

            if (_data.SmtpConfig.MaxSendsPerDay > 0)
            {
                var todaySent = _data.SentRecords.Count(r => r.Success && r.FechaEnvio.Date == DateTime.Today);
                var remainingToday = Math.Max(0, _data.SmtpConfig.MaxSendsPerDay - todaySent);
                result = result.Take(remainingToday).ToList();
            }

            if (result.Count < rows.Count)
            {
                var sessionLimit = _data.SmtpConfig.MaxSendsPerSession > 0 ? _data.SmtpConfig.MaxSendsPerSession.ToString() : "∞";
                var dayLimit = _data.SmtpConfig.MaxSendsPerDay > 0 ? _data.SmtpConfig.MaxSendsPerDay.ToString() : "∞";
                StatusMessage = $"Aplicando límites de envío ({scope}): sesión={sessionLimit}, día={dayLimit}.";
            }

            return result;
        }

        private List<StudentEntry> ApplyStudentSendLimits(List<StudentEntry> rows, string scope)
        {
            var result = rows;

            if (_data.SmtpConfig.MaxSendsPerSession > 0)
                result = result.Take(_data.SmtpConfig.MaxSendsPerSession).ToList();

            if (_data.SmtpConfig.MaxSendsPerDay > 0)
            {
                var todaySent = _data.SentRecords.Count(r => r.Success && r.FechaEnvio.Date == DateTime.Today);
                var remainingToday = Math.Max(0, _data.SmtpConfig.MaxSendsPerDay - todaySent);
                result = result.Take(remainingToday).ToList();
            }

            if (result.Count < rows.Count)
            {
                var sessionLimit = _data.SmtpConfig.MaxSendsPerSession > 0 ? _data.SmtpConfig.MaxSendsPerSession.ToString() : "∞";
                var dayLimit = _data.SmtpConfig.MaxSendsPerDay > 0 ? _data.SmtpConfig.MaxSendsPerDay.ToString() : "∞";
                StatusMessage = $"Aplicando límites de envío ({scope}): sesión={sessionLimit}, día={dayLimit}.";
            }

            return result;
        }

        private CourseInfo? GetCourseForStudent(StudentEntry student)
        {
            var raw = student.CursoRaw;
            if (!string.IsNullOrWhiteSpace(raw))
            {
                var byRaw = _data.GetCourseByRaw(raw);
                if (byRaw != null) return byRaw;
            }

            var display = student.CursoDisplay;
            return Courses.FirstOrDefault(c => string.Equals(c.NombreCorto, display, StringComparison.OrdinalIgnoreCase))
                ?? Courses.FirstOrDefault(c => string.Equals(c.CursoRaw?.Trim(), display?.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        private static Lead BuildLeadFromStudent(StudentEntry student)
        {
            if (student.LeadRow != null) return student.LeadRow.Lead;

            return new Lead
            {
                SourceRow = 0,
                FechaRegistro = student.FechaRegistro,
                Curso = student.CursoRaw,
                Plataforma = student.Plataforma,
                SituacionLaboral = student.SituacionLaboral,
                NivelEstudios = student.NivelEstudios,
                Nombre = student.Nombre,
                Email = student.Email,
                Telefono = student.Telefono,
                Provincia = student.Provincia,
                AceptaPublicidad = student.AceptaPublicidad,
                Observaciones = student.Observaciones,
                Contacto = student.Contacto,
                Inscripcion = student.Inscripcion
            };
        }

        private static LeadLabel GetStudentLabel(StudentEntry student)
            => student.LeadRow?.Label ?? LeadLabel.Confirmado;

        private string GetStudentLeadKey(StudentEntry student, Lead lead)
        {
            if (student.LeadRow != null) return _data.BuildLeadKey(lead);
            if (!string.IsNullOrWhiteSpace(student.Model?.Id)) return $"student:{student.Model.Id}";
            return _data.BuildLeadKey(lead);
        }

        private async Task<(bool Success, string? Error)> SendStudentWithCourseAsync(StudentEntry student, CourseInfo course)
        {
            var lead = BuildLeadFromStudent(student);
            if (!IsValidEmail(lead.Email))
            {
                var invalidError = "Email inválido.";
                _data.AddSentRecord(new SentRecord
                {
                    LeadKey = GetStudentLeadKey(student, lead),
                    NombreLead = lead.Nombre,
                    Email = lead.Email,
                    CursoRaw = string.IsNullOrWhiteSpace(course.CursoRaw) ? lead.Curso : course.CursoRaw,
                    FechaEnvio = DateTime.Now,
                    Success = false,
                    Error = invalidError
                });
                RefreshHistory();
                return (false, invalidError);
            }

            var labeledCourse = ApplyLabelTemplateToCourse(course, GetStudentLabel(student));
            var effectiveCourse = ApplyStartMessageToCourse(labeledCourse, student);
            var (success, error) = await _emailService.SendAsync(lead, effectiveCourse, _data.SmtpConfig);
            _data.AddSentRecord(new SentRecord
            {
                LeadKey = GetStudentLeadKey(student, lead),
                NombreLead = lead.Nombre,
                Email = lead.Email,
                CursoRaw = string.IsNullOrWhiteSpace(course.CursoRaw) ? lead.Curso : course.CursoRaw,
                FechaEnvio = DateTime.Now,
                Success = success,
                Error = error
            });
            RefreshHistory();
            return (success, error);
        }

        private CourseInfo ApplyStartMessageToCourse(CourseInfo course, StudentEntry student)
        {
            var template = (course.TextoInicioCurso ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(template)) return course;

            var courseName = string.IsNullOrWhiteSpace(student.CursoDisplay) ? course.NombreCorto : student.CursoDisplay;
            var message = template
                .Replace("{nombre}", student.Nombre)
                .Replace("{curso}", courseName);

            var clone = Clone(course);
            clone.TextoMarketing = string.IsNullOrWhiteSpace(clone.TextoMarketing)
                ? message
                : $"{message}\n\n{clone.TextoMarketing}";
            clone.TextoMarketingRich = RichTextSerialization.PlainTextToXaml(clone.TextoMarketing);

            return clone;
        }

        private string GetLabelTemplate(LeadLabel label)
            => label switch
            {
                LeadLabel.Solicitado => _data.SmtpConfig.LabelTemplateInteresado ?? string.Empty,
                LeadLabel.ConfirmaSi => _data.SmtpConfig.LabelTemplateConfirmado ?? string.Empty,
                _ when label.IsDiscardLabel() => _data.SmtpConfig.LabelTemplateDescartado ?? string.Empty,
                _ => string.Empty
            };

        private CourseInfo ApplyLabelTemplateToCourse(CourseInfo course, LeadLabel label)
        {
            var template = GetLabelTemplate(label).Trim();
            if (string.IsNullOrWhiteSpace(template)) return course;

            var clone = Clone(course);
            clone.TextoMarketing = string.IsNullOrWhiteSpace(clone.TextoMarketing)
                ? template
                : $"{template}\n\n{clone.TextoMarketing}";
            clone.TextoMarketingRich = RichTextSerialization.PlainTextToXaml(clone.TextoMarketing);

            return clone;
        }

        private static CourseInfo Clone(CourseInfo s) => new()
        {
            Id = s.Id, CursoRaw = s.CursoRaw, NombreCorto = s.NombreCorto, NombreComplementario = s.NombreComplementario,
            AsuntoEmail = s.AsuntoEmail, TextoMarketing = s.TextoMarketing, TextoMarketingRich = s.TextoMarketingRich,
            TextoInicioCurso = s.TextoInicioCurso, TextoWhatsApp = s.TextoWhatsApp, TextoSocial = s.TextoSocial,
            RequisitosAcceso = s.RequisitosAcceso, DocumentacionNecesaria = s.DocumentacionNecesaria,
            FechaInicio = s.FechaInicio, FechaFin = s.FechaFin, HorarioInfo = s.HorarioInfo,
            UrlFichaInscripcion = s.UrlFichaInscripcion, PdfAdjuntoPath = s.PdfAdjuntoPath,
            InfoAdicional = s.InfoAdicional, FechaCreacion = s.FechaCreacion, FechaModificacion = s.FechaModificacion
        };

        private async Task<(bool Success, string? Error)> SendLeadAsync(LeadRow row, bool forceSend)
        {
            var course = _data.GetCourseByRaw(row.Lead.Curso);
            if (course == null)
            {
                var err = "No se encontró la configuración del curso para este lead.";
                _data.AddSentRecord(new SentRecord
                {
                    LeadKey    = _data.BuildLeadKey(row.Lead),
                    NombreLead = row.Lead.Nombre,
                    Email      = row.Lead.Email,
                    CursoRaw   = row.Lead.Curso,
                    FechaEnvio = DateTime.Now,
                    Success    = false,
                    Error      = err
                });
                return (false, err);
            }

            return await SendLeadWithCourseAsync(row, course, forceSend, row.Lead.Curso);
        }

        private void MarkLeadEmailSent(LeadRow row, DateTime sentAt)
        {
            var key = _data.BuildLeadKey(row.Lead);
            row.EmailSentAt = sentAt;
            _data.SetLeadEmailSentAt(key, sentAt);
            MarkLeadHandled(row);
            UpdateCounts();
            RefreshFilter();
        }

        private void MarkLeadWhatsAppOpened(LeadRow row, DateTime sentAt)
        {
            var key = _data.BuildLeadKey(row.Lead);
            row.WhatsAppSentAt = sentAt;
            _data.SetLeadWhatsAppSentAt(key, sentAt);
            MarkLeadHandled(row);
            UpdateCounts();
            RefreshFilter();
        }

        private void MarkLeadHandled(LeadRow row)
        {
            var handledBy = GetCurrentCaptadoraName();
            if (string.IsNullOrWhiteSpace(handledBy)) return;

            var key = _data.BuildLeadKey(row.Lead);
            row.HandledBy = handledBy;
            _data.SetLeadHandledBy(key, handledBy);
        }

        private string GetCurrentCaptadoraName()
        {
            var nombre = (SmtpConfig.CaptadoraNombre ?? string.Empty).Trim();
            var apellidos = (SmtpConfig.CaptadoraApellidos ?? string.Empty).Trim();
            return string.Join(" ", new[] { nombre, apellidos }.Where(x => !string.IsNullOrWhiteSpace(x))).Trim();
        }

        private async Task<(bool Success, string? Error)> SendLeadWithCourseAsync(
            LeadRow row,
            CourseInfo course,
            bool forceSend,
            string? recordCourseRaw)
        {
            if (!forceSend && (row.AlreadySent || row.IsDiscarded))
                return (false, "Lead no elegible para envío.");

            if (!IsValidEmail(row.Email))
            {
                var invalidError = "Email inválido.";
                _data.AddSentRecord(new SentRecord
                {
                    LeadKey = _data.BuildLeadKey(row.Lead),
                    NombreLead = row.Lead.Nombre,
                    Email = row.Lead.Email,
                    CursoRaw = string.IsNullOrWhiteSpace(recordCourseRaw) ? row.Lead.Curso : recordCourseRaw,
                    FechaEnvio = DateTime.Now,
                    Success = false,
                    Error = invalidError
                });
                return (false, invalidError);
            }

            var effectiveCourse = ApplyLabelTemplateToCourse(course, row.Label);
            var (success, error) = await _emailService.SendAsync(row.Lead, effectiveCourse, _data.SmtpConfig);
            var sentAt = DateTime.Now;
            _data.AddSentRecord(new SentRecord
            {
                LeadKey    = _data.BuildLeadKey(row.Lead),
                NombreLead = row.Lead.Nombre,
                Email      = row.Lead.Email,
                CursoRaw   = string.IsNullOrWhiteSpace(recordCourseRaw) ? row.Lead.Curso : recordCourseRaw,
                FechaEnvio = sentAt,
                Success    = success,
                Error      = error
            });

            if (success) MarkLeadEmailSent(row, sentAt);
            RefreshHistory();
            return (success, error);
        }

        private static bool IsValidEmail(string email)
        {
            try
            {
                _ = new MailAddress(email);
                return true;
            }
            catch { return false; }
        }

        private static string Csv(string value)
        {
            value ??= string.Empty;
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }

        private string GetLeadLabelDisplay(LeadLabel label)
            => LeadLabelOptions.FirstOrDefault(x => x.Value == label)?.Display ?? label.ToString();

        private static string NormalizePhoneForWhatsApp(string phone)
        {
            if (string.IsNullOrWhiteSpace(phone)) return string.Empty;
            
            // Remove all non-digit characters
            var digits = new string(phone.Where(char.IsDigit).ToArray());
            
            if (string.IsNullOrEmpty(digits)) return string.Empty;
            
            // If starts with 00, remove it (international prefix)
            if (digits.StartsWith("00"))
                digits = digits.Substring(2);
            
            // If doesn't start with a country code (assuming 9-digit numbers need country code)
            // Add Spain country code (34) as default
            if (digits.Length == 9)
                digits = "34" + digits;
            
            return digits;
        }

        private static void CreateAndOpenVCard(string nombre, string telefono, string? email = null)
        {
            try
            {
                var tempDir = Path.Combine(Path.GetTempPath(), "LeadMailer");
                Directory.CreateDirectory(tempDir);
                
                // Crear archivo vCard con nombre único
                var vCardPath = Path.Combine(tempDir, $"{nombre.Replace(" ", "_")}_{Guid.NewGuid().ToString().Substring(0, 8)}.vcf");
                
                var vCard = new StringBuilder();
                vCard.AppendLine("BEGIN:VCARD");
                vCard.AppendLine("VERSION:3.0");
                vCard.AppendLine($"FN:{EscapeVCardValue(nombre)}");
                vCard.AppendLine($"TEL:{telefono}");
                
                if (!string.IsNullOrWhiteSpace(email))
                    vCard.AppendLine($"EMAIL:{EscapeVCardValue(email)}");
                
                vCard.AppendLine("END:VCARD");
                
                File.WriteAllText(vCardPath, vCard.ToString(), Encoding.UTF8);
                
                // Abrir el archivo vCard para que Windows lo agregue a contactos
                Process.Start(new ProcessStartInfo
                {
                    FileName = vCardPath,
                    UseShellExecute = true
                });
                
                // Esperar a que el usuario agregue el contacto antes de abrir WhatsApp
                System.Threading.Thread.Sleep(1500);
            }
            catch (Exception ex)
            {
                // El WhatsApp se abrirá igualmente sin el contacto
                Services.AppLogger.Warn($"CreateAndOpenVCard: no se pudo crear el contacto para '{nombre}' — {ex.Message}");
            }
        }

        private static string EscapeVCardValue(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            
            // vCard escaping: escape commas y saltos de línea
            return value.Replace(";", "\\;").Replace(",", "\\,").Replace("\n", "\\n");
        }

        [RelayCommand]
        private void SendWhatsAppToSelected()
        {
            var rows = Leads.Where(l => l.IsSelected).ToList();
            if (rows.Count == 0)
            {
                StatusMessage = "Selecciona al menos un lead para WhatsApp.";
                return;
            }

            int opened = 0, invalid = 0;
            foreach (var row in rows)
            {
                if (TryOpenWhatsApp(row, row.Course))
                {
                    opened++;
                    MarkLeadWhatsAppOpened(row, DateTime.Now);
                }
                else invalid++;
            }

            StatusMessage = invalid > 0
                ? $"WhatsApp: {opened} chats abiertos, {invalid} sin teléfono válido."
                : $"WhatsApp: {opened} chats abiertos.";
        }

        [RelayCommand]
        private void PreviewLeadEmail(LeadRow? row)
        {
            if (row == null) return;
            ShowEmailPreview(row, row.Course);
        }

        [RelayCommand]
        private void PreviewSelectedEmail()
        {
            var row = Leads.FirstOrDefault(l => l.IsSelected) ?? Leads.FirstOrDefault();
            if (row == null)
            {
                StatusMessage = "No hay leads para previsualizar.";
                return;
            }

            ShowEmailPreview(row, row.Course);
        }

        [RelayCommand]
        private void CopySelectedEmails()
        {
            var emails = Leads
                .Where(l => l.IsSelected)
                .Select(l => (l.Email ?? string.Empty).Trim())
                .Where(e => !string.IsNullOrWhiteSpace(e))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (emails.Count == 0)
            {
                StatusMessage = "No hay emails seleccionados para copiar.";
                return;
            }

            try
            {
                Clipboard.SetText(string.Join(";", emails));
                StatusMessage = $"✓  Copiados {emails.Count} emails al portapapeles.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"✗  No se pudo copiar al portapapeles: {ex.Message}";
            }
        }

        [RelayCommand]
        private void SendWhatsApp(LeadRow? row)
        {
            if (row == null) return;

            if (TryOpenWhatsApp(row, row.Course))
            {
                MarkLeadWhatsAppOpened(row, DateTime.Now);
                StatusMessage = $"WhatsApp abierto para {row.Nombre}.";
            }
            else
                StatusMessage = $"No se pudo abrir WhatsApp para {row.Nombre}: teléfono no válido.";
        }

        private void ShowEmailPreview(LeadRow row, CourseInfo course)
        {
            try
            {
                var effectiveCourse = ApplyLabelTemplateToCourse(course, row.Label);
                var preview = _emailService.BuildPreview(row.Lead, effectiveCourse, _data.SmtpConfig);
                var tempFile = Path.Combine(Path.GetTempPath(), $"leadmailer_preview_{Guid.NewGuid():N}.html");
                File.WriteAllText(tempFile, preview.HtmlBody, Encoding.UTF8);

                Process.Start(new ProcessStartInfo
                {
                    FileName = tempFile,
                    UseShellExecute = true
                });

                StatusMessage = $"👁  Previsualización abierta para {row.Email}.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"✗  No se pudo generar la previsualización: {ex.Message}";
            }
        }

        private static bool TryOpenWhatsApp(StudentEntry student, CourseInfo course)
        {
            var phone = NormalizePhoneForWhatsApp(student.Telefono);
            if (string.IsNullOrWhiteSpace(phone)) return false;

            CreateAndOpenVCard(student.Nombre, phone, student.Email);

            var courseName = string.IsNullOrWhiteSpace(student.CursoDisplay) ? course.NombreCorto : student.CursoDisplay;
            var startTemplate = (course.TextoInicioCurso ?? string.Empty).Trim();
            var baseTemplate = !string.IsNullOrWhiteSpace(startTemplate)
                ? startTemplate
                : course.TextoWhatsApp;
            var text = !string.IsNullOrWhiteSpace(baseTemplate)
                ? baseTemplate
                    .Replace("{nombre}", student.Nombre)
                    .Replace("{curso}", courseName)
                : $"Hola {student.Nombre}, te escribimos por tu interés en el curso \"{courseName}\".";

            try
            {
                var url = $"whatsapp://send?phone={phone}&text={Uri.EscapeDataString(text)}";
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
                return true;
            }
            catch (Exception ex)
            {
                Services.AppLogger.Warn($"WhatsApp desktop no disponible para {phone}, intentando web: {ex.Message}");
                try
                {
                    var url = $"https://wa.me/{phone}?text={Uri.EscapeDataString(text)}";
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = url,
                        UseShellExecute = true
                    });
                    return true;
                }
                catch (Exception ex2)
                {
                    Services.AppLogger.Error($"TryOpenWhatsApp: no se pudo abrir WhatsApp web para {phone}", ex2);
                    return false;
                }
            }
        }

        private static bool TryOpenWhatsApp(LeadRow row, CourseInfo course)
        {
            var phone = NormalizePhoneForWhatsApp(row.Telefono);
            if (string.IsNullOrWhiteSpace(phone)) return false;

            // Crear contacto vCard y agregarlo a la agenda
            CreateAndOpenVCard(row.Nombre, phone, row.Email);

            // Usar el mensaje personalizado del curso si está configurado
            var text = !string.IsNullOrWhiteSpace(course.TextoWhatsApp)
                ? course.TextoWhatsApp
                    .Replace("{nombre}", row.Nombre)
                    .Replace("{curso}", row.NombreCurso)
                : $"Hola {row.Nombre}, te escribimos por tu interés en el curso \"{row.NombreCurso}\".";
            
            try
            {
                var url = $"whatsapp://send?phone={phone}&text={Uri.EscapeDataString(text)}";
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
                return true;
            }
            catch (Exception ex)
            {
                Services.AppLogger.Warn($"WhatsApp desktop no disponible para {phone}, intentando web: {ex.Message}");
                try
                {
                    var url = $"https://wa.me/{phone}?text={Uri.EscapeDataString(text)}";
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = url,
                        UseShellExecute = true
                    });
                    return true;
                }
                catch (Exception ex2)
                {
                    Services.AppLogger.Error($"TryOpenWhatsApp: no se pudo abrir WhatsApp web para {phone}", ex2);
                    return false;
                }
            }
        }
    }
}
