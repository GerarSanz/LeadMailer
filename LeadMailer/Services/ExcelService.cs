using System.IO;
using System.Globalization;
using System.Text;
using LeadMailer.Models;
using OfficeOpenXml;

namespace LeadMailer.Services;

/// <summary>Lee leads desde un archivo Excel (.xlsx / .xls).</summary>
public class ExcelService
{
    private static bool _licenseConfigured;

    private static void EnsureEpplusLicense()
    {
        if (_licenseConfigured) return;

        ExcelPackage.License.SetNonCommercialPersonal("LeadMailer");

        _licenseConfigured = true;
    }

    public List<Lead> ReadLeads(string filePath)
    {
        EnsureEpplusLicense();
        var leads = new List<Lead>();

        if (!File.Exists(filePath))
            throw new FileNotFoundException($"No se encontró el archivo: {filePath}");

        using var package = new ExcelPackage(new FileInfo(filePath));

        static string NormalizeHeader(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;

            var formD = value.Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD);
            var chars = formD.Where(ch => CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark);
            var noAccents = new string(chars.ToArray()).Normalize(NormalizationForm.FormC);

            return new string(noAccents.Where(ch => char.IsLetterOrDigit(ch)).ToArray());
        }

        static int ScoreWorksheet(ExcelWorksheet sheet)
        {
            if (sheet.Dimension == null) return int.MinValue;

            var maxRow = Math.Min(sheet.Dimension.End.Row, 10);
            var maxCol = Math.Min(sheet.Dimension.End.Column, 80);
            var score = 0;

            for (int row = 1; row <= maxRow; row++)
            {
                for (int c = 1; c <= maxCol; c++)
                {
                    var h = NormalizeHeader(sheet.Cells[row, c].Text);
                    if (string.IsNullOrWhiteSpace(h)) continue;

                    if (h == "curso") score += 15;
                    else if (h.Contains("curso")) score += 8;

                    if (h.Contains("email") || h.Contains("correo")) score += 3;
                    if (h.Contains("nombre")) score += 2;
                    if (h.Contains("telefono")) score += 1;
                }
            }

            return score;
        }

        var ws = package.Workbook.Worksheets
            .Where(s => s?.Dimension != null)
            .OrderByDescending(ScoreWorksheet)
            .FirstOrDefault();

        if (ws?.Dimension == null) return leads;

        int GuessHeaderRow()
        {
            var maxRow = Math.Min(ws.Dimension.End.Row, 10);
            var probes = new[] { "curso", "cursos", "formacion", "accion", "email", "correo", "nombre", "telefono", "provincia", "plataforma" };

            var bestRow = 1;
            var bestScore = -1;
            var bestNonEmpty = -1;

            for (int row = 1; row <= maxRow; row++)
            {
                var score = 0;
                var nonEmpty = 0;
                var hasCursoHeader = false;
                for (int c = 1; c <= ws.Dimension.End.Column; c++)
                {
                    var normalized = NormalizeHeader(ws.Cells[row, c].Text);
                    if (string.IsNullOrWhiteSpace(normalized)) continue;
                    nonEmpty++;
                    if (normalized == "curso" || normalized.Contains("curso"))
                        hasCursoHeader = true;
                    if (probes.Any(p => normalized.Contains(p))) score++;
                }

                if (hasCursoHeader)
                    score += 5;

                if (score > bestScore || (score == bestScore && nonEmpty > bestNonEmpty))
                {
                    bestScore = score;
                    bestNonEmpty = nonEmpty;
                    bestRow = row;
                }
            }

            return bestRow;
        }

        var headerRow = GuessHeaderRow();

        // ── Mapear columnas por cabecera ──────────────────────────────────────
        var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int c = 1; c <= ws.Dimension.End.Column; c++)
        {
            var raw = ws.Cells[headerRow, c].Text.Trim();
            if (string.IsNullOrWhiteSpace(raw)) continue;

            var normalized = NormalizeHeader(raw);
            if (!colMap.ContainsKey(raw)) colMap[raw] = c;
            if (!string.IsNullOrWhiteSpace(normalized) && !colMap.ContainsKey(normalized)) colMap[normalized] = c;
        }

        int Col(params string[] names)
        {
            foreach (var n in names)
            {
                if (colMap.TryGetValue(n, out var col)) return col;

                var normalized = NormalizeHeader(n);
                if (!string.IsNullOrWhiteSpace(normalized) && colMap.TryGetValue(normalized, out col))
                    return col;
            }
            return -1;
        }

        int ColContains(string fragment)
        {
            foreach (var kv in colMap)
                if (kv.Key.Contains(fragment, StringComparison.OrdinalIgnoreCase)
                    || kv.Key.Contains(NormalizeHeader(fragment), StringComparison.OrdinalIgnoreCase))
                    return kv.Value;
            return -1;
        }

        int cFecha    = Col("  ", "Fecha", "Timestamp");
        int cCurso    = Col(
            "curso", "cursos", "cursointeres", "cursointeresado", "cursoalqueseapunto",
            "formacion", "accionformativa", "accinformativa", "accformativa");
        if (cCurso < 0)
            cCurso = ColContains("curso");
        if (cCurso < 0)
            cCurso = ColContains("formacion");
        if (cCurso < 0)
            cCurso = ColContains("accion");
        int cPlat     = Col("Plataforma");
        int cSit      = Col("Situación Laboral", "Situacion Laboral");
        int cNivel    = Col("Sector Laboral / Nivel Estudios", "Nivel Estudios");
        int cNombre   = Col("Nombre");
        int cEmail    = Col("Email", "Correo", "E-mail");
        int cTel      = Col("Teléfono", "Telefono");
        int cProv     = Col("Provincia");
        int cPub      = Col("Aceptación Publicidad", "Aceptacion Publicidad");
        int cObs      = Col("OBSERVACIONES", "Observaciones");
        int cContact  = Col("CONTACTO. PERSONA / FECHA / MEDIO / RESULTADO", "Contacto");
        int cInsc     = Col("INSCRIPCIÓN (SI/NO)", "Inscripción", "Inscripcion");

        string Cell(int row, int col) => col > 0 ? ws.Cells[row, col].Text.Trim() : "";
        
        // Lectura especial para teléfono: asegurar que se lee como texto puro
        string CellAsText(int row, int col)
        {
            if (col <= 0) return "";
            var cell = ws.Cells[row, col];
            // Intentar leer como string primero
            if (cell.Value is string str)
                return str.Trim();
            // Si es número, convertir a string sin formato científico
            if (cell.Value is double d)
                return d.ToString("0");
            if (cell.Value is long l)
                return l.ToString();
            if (cell.Value is int i)
                return i.ToString();
            // Fallback al texto normal
            return cell.Text.Trim();
        }

        // ── Leer filas ────────────────────────────────────────────────────────
        for (int row = headerRow + 1; row <= ws.Dimension.End.Row; row++)
        {
            var email = Cell(row, cEmail);
            var curso = Cell(row, cCurso);
            if (string.IsNullOrWhiteSpace(email) && string.IsNullOrWhiteSpace(curso)) continue;

            var pubText = Cell(row, cPub).ToLowerInvariant();
            bool acepta = pubText is "true" or "1" or "sí" or "si" or "yes";

            leads.Add(new Lead
            {
                SourceRow        = row,
                FechaRegistro    = Cell(row, cFecha),
                Curso            = curso,
                Plataforma       = Cell(row, cPlat),
                SituacionLaboral = Cell(row, cSit),
                NivelEstudios    = Cell(row, cNivel),
                Nombre           = Cell(row, cNombre),
                Email            = email,
                Telefono         = CellAsText(row, cTel),  // Usar CellAsText para evitar formato científico
                Provincia        = Cell(row, cProv),
                AceptaPublicidad = acepta,
                Observaciones    = Cell(row, cObs),
                Contacto         = Cell(row, cContact),
                Inscripcion      = Cell(row, cInsc)
            });
        }

        return leads;
    }
}
