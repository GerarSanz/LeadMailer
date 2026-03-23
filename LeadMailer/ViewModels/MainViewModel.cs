using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using LeadMailer.Helpers;
using LeadMailer.Models;
using LeadMailer.Services;
using Microsoft.Win32;
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
        [ObservableProperty] private string _whatsappMessage = "";

        public bool SmtpTestResultIsError => SmtpTestResult.StartsWith("✗");

        public string AppVersion
        {
            get
            {
                var v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                return v is null ? "v?" : $"v{v.Major}.{v.Minor}.{v.Build}";
            }
        }

        public IReadOnlyList<string> StatusFilterOptions { get; } =
            new[] { "Todos", "Pendiente", "Enviado", "Interesado", "Confirmado", "Descartado", "Seguimiento hoy" };

        public IReadOnlyList<LabelOption> LeadLabelOptions { get; } = new LabelOption[]
        {
            new(LeadLabel.Ninguna,    "—"),
            new(LeadLabel.Interesado, "Interesado"),
            new(LeadLabel.Confirmado, "Confirmado"),
            new(LeadLabel.Descartado, "Descartado"),
        };

        public IReadOnlyList<string> HistoryStatusFilterOptions { get; } =
            new[] { "Todos", "Correctos", "Fallidos" };

        private ICollectionView? _filteredLeads;
        public ICollectionView FilteredLeads => _filteredLeads!;

        private ICollectionView? _filteredHistory;
        public ICollectionView FilteredHistory => _filteredHistory!;

        private readonly List<string> _lastFailedLeadKeys = new();

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
            RefreshHistory();
        }

        // ── Reacción a cambios de filtro ──────────────────────────────────────
        partial void OnSearchTextChanged(string value)              => RefreshFilter();
        partial void OnSelectedCourseFilterChanged(CourseInfo value) => RefreshFilter();
        partial void OnSelectedStatusFilterChanged(string value)    => RefreshFilter();
        partial void OnHistorySearchTextChanged(string value)       => RefreshHistoryFilter();
        partial void OnHistoryStatusFilterChanged(string value)     => RefreshHistoryFilter();

        [RelayCommand] private void Navigate(string view) => CurrentView = view;

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
                    bool alreadySent = _data.HasBeenSent(lead);
                    var  key         = _data.BuildLeadKey(lead);
                    var  label       = _data.GetLeadLabel(key);
                    var note         = _data.GetLeadNote(key);
                    var nextContact  = _data.GetLeadNextContact(key);

                    if (!IsCourseConfigured(course) && !newCourses.Contains(lead.Curso))
                        newCourses.Add(lead.Curso);

                    Leads.Add(new LeadRow(
                        lead: lead,
                        course: course,
                        alreadySent: alreadySent,
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
                            UpdateCounts();
                            RefreshFilter();
                        }));
                }
                RefreshCourses();
                UpdateCounts();
                RefreshFilter();

                UpdateCourseWarning();
                StatusMessage = rawLeads.Count == 0
                    ? "El archivo no contiene filas de datos."
                    : $"✓  {rawLeads.Count} leads cargados — {PendingCount} pendientes de envío.";
            }
            catch (Exception ex) { StatusMessage = $"Error al leer el Excel: {ex.Message}"; }
        }

        [RelayCommand]
        private async Task SendSelected()
        {
            var pending = Leads.Where(l => l.IsSelected && !l.AlreadySent && !l.IsDiscarded).ToList();
            if (pending.Count == 0) { StatusMessage = "No hay leads seleccionados sin enviar."; return; }
            if (!SmtpOk()) return;

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
        }

        [RelayCommand] private void SelectAll()   { foreach (var l in Leads.Where(x => x.CanSelect)) l.IsSelected = true; }
        [RelayCommand] private void DeselectAll() { foreach (var l in Leads) l.IsSelected = false; }

        [RelayCommand]
        private void SelectFollowUpToday()
        {
            foreach (var l in Leads)
                l.IsSelected = l.CanSelect && IsFollowUpToday(l.NextContact);

            var count = Leads.Count(l => l.IsSelected);
            StatusMessage = count == 0
                ? "No hay leads con seguimiento para hoy."
                : $"✓  Seleccionados {count} leads con seguimiento de hoy.";
        }

        [RelayCommand]
        private void ExportFilteredLeads()
        {
            var rows = FilteredLeads?.Cast<LeadRow>().ToList() ?? new List<LeadRow>();
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
            sb.AppendLine("Estado,Nombre,Email,Curso,Plataforma,Provincia,SituacionLaboral,Telefono,NotaSeguimiento,ProximoContacto");
            foreach (var r in rows)
            {
                sb.AppendLine(string.Join(",",
                    Csv(r.EstadoTexto), Csv(r.Nombre), Csv(r.Email), Csv(r.NombreCurso),
                    Csv(r.Plataforma), Csv(r.Provincia), Csv(r.SituacionLaboral), Csv(r.Telefono),
                    Csv(r.Note), Csv(r.NextContact)));
            }

            File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
            StatusMessage = $"✓  Exportados {rows.Count} leads a CSV.";
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
            row.Label = row.Label == LeadLabel.Descartado ? LeadLabel.Ninguna : LeadLabel.Descartado;
            // persistencia y refresco gestionados por el callback onLabelChanged
        }

        [RelayCommand]
        private void DeleteLead(LeadRow? row)
        {
            if (row == null) return;
            Leads.Remove(row);
            var key = _data.BuildLeadKey(row.Lead);
            _data.RemoveLeadLabel(key);
            _data.RemoveLeadNote(key);
            _data.RemoveLeadNextContact(key);
            UpdateCounts();
            RefreshFilter();
            StatusMessage = $"✓  Lead '{row.Nombre}' eliminado de la lista.";
        }

        // ── Cursos ────────────────────────────────────────────────────────────

        [RelayCommand]
        private void EditCourse(CourseInfo? course)
        {
            if (course == null) return;
            EditingCourse = Clone(course);
            WhatsappMessage = course.TextoWhatsApp ?? "";
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

        [RelayCommand] private void CancelEdit() { EditingCourse = null; WhatsappMessage = ""; CurrentView = "Courses"; }

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
            _data.Save();
            SmtpTestResult = "✓  Configuración guardada.";
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
                    "Enviado"          => row.Status == LeadStatus.Enviado,
                    "Interesado"       => row.Label  == LeadLabel.Interesado,
                    "Confirmado"       => row.Label  == LeadLabel.Confirmado,
                    "Descartado"       => row.Status == LeadStatus.Descartado,
                    "Seguimiento hoy"  => IsFollowUpToday(row.NextContact),
                    _                  => true
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
        }

        private void RefreshHistoryFilter()
        {
            _filteredHistory?.Refresh();
        }

        private void RefreshHistory()
        {
            SentHistory.Clear();
            foreach (var r in _data.SentRecords.OrderByDescending(x => x.FechaEnvio))
                SentHistory.Add(r);
            RefreshHistoryFilter();
        }

        private void RefreshCourses()
        {
            Courses.Clear();
            var previousId = SelectedCourseFilter?.Id;

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
        }

        private void UpdateCounts()
        {
            PendingCount      = Leads.Count(l => l.Status == LeadStatus.Pendiente);
            SentCount         = Leads.Count(l => l.Status == LeadStatus.Enviado);
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

        private static CourseInfo Clone(CourseInfo s) => new()
        {
            Id = s.Id, CursoRaw = s.CursoRaw, NombreCorto = s.NombreCorto, NombreComplementario = s.NombreComplementario,
            AsuntoEmail = s.AsuntoEmail, TextoMarketing = s.TextoMarketing, TextoMarketingRich = s.TextoMarketingRich,
            TextoWhatsApp = s.TextoWhatsApp,
            RequisitosAcceso = s.RequisitosAcceso, DocumentacionNecesaria = s.DocumentacionNecesaria,
            FechaInicio = s.FechaInicio, FechaFin = s.FechaFin, HorarioInfo = s.HorarioInfo,
            UrlFichaInscripcion = s.UrlFichaInscripcion, PdfAdjuntoPath = s.PdfAdjuntoPath,
            InfoAdicional = s.InfoAdicional, FechaCreacion = s.FechaCreacion, FechaModificacion = s.FechaModificacion
        };

        private async Task<(bool Success, string? Error)> SendLeadAsync(LeadRow row, bool forceSend)
        {
            if (!forceSend && (row.AlreadySent || row.IsDiscarded))
                return (false, "Lead no elegible para envío.");

            if (!IsValidEmail(row.Email))
            {
                var invalidError = "Email inválido.";
                _data.AddSentRecord(new SentRecord
                {
                    LeadKey    = _data.BuildLeadKey(row.Lead),
                    NombreLead = row.Lead.Nombre,
                    Email      = row.Lead.Email,
                    CursoRaw   = row.Lead.Curso,
                    FechaEnvio = DateTime.Now,
                    Success    = false,
                    Error      = invalidError
                });
                return (false, invalidError);
            }

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

            var (success, error) = await _emailService.SendAsync(row.Lead, course, _data.SmtpConfig);
            _data.AddSentRecord(new SentRecord
            {
                LeadKey    = _data.BuildLeadKey(row.Lead),
                NombreLead = row.Lead.Nombre,
                Email      = row.Lead.Email,
                CursoRaw   = row.Lead.Curso,
                FechaEnvio = DateTime.Now,
                Success    = success,
                Error      = error
            });

            if (success) row.AlreadySent = true;
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
                if (TryOpenWhatsApp(row, row.Course)) opened++;
                else invalid++;
            }

            StatusMessage = invalid > 0
                ? $"WhatsApp: {opened} chats abiertos, {invalid} sin teléfono válido."
                : $"WhatsApp: {opened} chats abiertos.";
        }

        [RelayCommand]
        private void SendWhatsApp(LeadRow? row)
        {
            if (row == null) return;

            if (TryOpenWhatsApp(row, row.Course))
                StatusMessage = $"WhatsApp abierto para {row.Nombre}.";
            else
                StatusMessage = $"No se pudo abrir WhatsApp para {row.Nombre}: teléfono no válido.";
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
